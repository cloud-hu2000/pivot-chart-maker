import Papa from "papaparse";
import * as XLSX from "@e965/xlsx";
import "./tool-pages.css";

const page = document.body.dataset.tool;
const $ = (selector) => document.querySelector(selector);
const status = (message, error = false) => {
  const target = $("#tool-status");
  target.textContent = message;
  target.classList.toggle("error", error);
};
const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]);
const prettySize = (bytes) => bytes < 1024 ? `${bytes} B` : bytes < 1048576 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1048576).toFixed(1)} MB`;
function download(name, data, type) {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
function uniqueHeaders(raw) {
  const used = new Set();
  return raw.map((value, index) => {
    const base = String(value ?? "").trim() || `Column ${index + 1}`;
    let name = base;
    let suffix = 2;
    while (used.has(name.toLowerCase())) name = `${base} ${suffix++}`;
    used.add(name.toLowerCase());
    return name;
  });
}
function rectangular(rows) {
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const headers = uniqueHeaders(Array.from({ length: width }, (_, index) => rows[0]?.[index]));
  return { headers, rows: rows.slice(1).map((row) => Array.from({ length: width }, (_, index) => row[index] ?? "")) };
}
function parseCsv(text, delimiter = "") {
  const result = Papa.parse(text.replace(/^\uFEFF/, ""), { skipEmptyLines: "greedy", delimiter: delimiter || undefined });
  if (result.errors.length) throw new Error(`CSV line ${result.errors[0].row + 1}: ${result.errors[0].message}`);
  if (!result.data.length) throw new Error("The file has no rows.");
  return rectangular(result.data);
}
async function decodeFile(file, encoding) {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  const detected = encoding || (bytes[0] === 0xff && bytes[1] === 0xfe ? "utf-16le" : "utf-8");
  return new TextDecoder(detected, { fatal: false }).decode(buffer).replace(/^\uFEFF/, "");
}
const delimiter = () => $("#delimiter")?.value || "";
const encoding = () => $("#encoding")?.value || "";
const fileName = (file) => file?.name.replace(/\.[^.]+$/, "") || "converted";
function tablePreview(headers, rows, maxRows = 12) {
  const box = $("#preview");
  if (!box) return;
  const shown = rows.slice(0, maxRows);
  box.innerHTML = `<table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${shown.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  box.hidden = false;
}
function flatten(value, prefix = "", output = {}) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    const entries = Object.entries(value);
    if (!entries.length && prefix) output[prefix] = "{}";
    for (const [key, child] of entries) flatten(child, prefix ? `${prefix}.${key}` : key, output);
  } else if (Array.isArray(value)) {
    output[prefix || "value"] = JSON.stringify(value);
  } else {
    output[prefix || "value"] = value == null ? "" : value;
  }
  return output;
}
function setupJsonToCsv() {
  let csv = "";
  let name = "converted";
  $("#convert").addEventListener("click", async () => {
    try {
      const file = $("#file").files[0];
      const source = file ? await file.text() : $("#source").value;
      if (!source.trim()) throw new Error("Paste JSON or select a .json file first.");
      const parsed = JSON.parse(source);
      const records = Array.isArray(parsed) ? parsed : [parsed];
      if (!records.length) throw new Error("The JSON array is empty.");
      const objects = records.map((record) => flatten(record));
      const headers = [...new Set(objects.flatMap(Object.keys))];
      const rows = objects.map((object) => headers.map((header) => object[header] ?? ""));
      csv = "\uFEFF" + Papa.unparse([headers, ...rows]);
      name = fileName(file);
      tablePreview(headers, rows);
      $("#download").disabled = false;
      status(`${rows.length.toLocaleString()} row${rows.length === 1 ? "" : "s"}, ${headers.length} columns ready. Nested keys use dot notation; arrays remain JSON text.`);
    } catch (error) { status(error.message, true); }
  });
  $("#download").addEventListener("click", () => download(`${name}.csv`, csv, "text/csv;charset=utf-8"));
}
function setupCsvToJson() {
  let json = "";
  let name = "converted";
  $("#convert").addEventListener("click", async () => {
    try {
      const file = $("#file").files[0];
      const source = file ? await decodeFile(file, encoding()) : $("#source").value;
      const { headers, rows } = parseCsv(source, delimiter());
      const objects = rows.map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
      json = JSON.stringify(objects, null, 2);
      name = fileName(file);
      $("#output").value = json;
      $("#download").disabled = false;
      status(`${objects.length.toLocaleString()} records ready. CSV values stay strings so IDs and leading zeros are preserved.`);
    } catch (error) { status(error.message, true); }
  });
  $("#download").addEventListener("click", () => download(`${name}.json`, json, "application/json;charset=utf-8"));
}
function setupCsvToExcel() {
  let workbook;
  let name = "converted";
  $("#convert").addEventListener("click", async () => {
    try {
      const file = $("#file").files[0];
      if (!file) throw new Error("Select a CSV file first.");
      const source = await decodeFile(file, encoding());
      const { headers, rows } = parseCsv(source, delimiter());
      if (headers.length > 16384 || rows.length > 1048575) throw new Error("This file exceeds Excel's sheet limits of 16,384 columns or 1,048,576 rows.");
      workbook = XLSX.utils.book_new();
      const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      sheet["!cols"] = headers.map((header) => ({ wch: Math.min(40, Math.max(12, header.length + 2)) }));
      XLSX.utils.book_append_sheet(workbook, sheet, "CSV data");
      name = fileName(file);
      tablePreview(headers, rows);
      $("#download").disabled = false;
      status(`${rows.length.toLocaleString()} rows ready. Cell values are kept as text to preserve leading zeros and exact IDs.`);
    } catch (error) { status(error.message, true); }
  });
  $("#download").addEventListener("click", () => {
    try { XLSX.writeFile(workbook, `${name}.xlsx`, { compression: true }); }
    catch (error) { status(error.message, true); }
  });
}
function setupViewer() {
  const fileInput = $("#file");
  const scroller = $("#table-scroll");
  const table = $("#viewer-table");
  const search = $("#search");
  const ROW_HEIGHT = 38;
  let headers = [];
  let rows = [];
  let matches = null;
  let currentFile;
  let searchTimer;
  let parsing = false;
  function render() {
    const total = matches ? matches.length : rows.length;
    const start = Math.max(0, Math.floor(scroller.scrollTop / ROW_HEIGHT) - 5);
    const count = Math.ceil(scroller.clientHeight / ROW_HEIGHT) + 12;
    const end = Math.min(total, start + count);
    let html = `<thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>`;
    html += `<tr class="spacer"><td colspan="${headers.length}" style="height:${start * ROW_HEIGHT}px"></td></tr>`;
    for (let i = start; i < end; i++) {
      const rowIndex = matches ? matches[i] : i;
      html += `<tr data-row="${rowIndex}">${rows[rowIndex].map((cell, column) => `<td contenteditable="true" data-col="${column}" title="Click to edit">${escapeHtml(cell)}</td>`).join("")}</tr>`;
    }
    html += `<tr class="spacer"><td colspan="${headers.length}" style="height:${Math.max(0, total - end) * ROW_HEIGHT}px"></td></tr></tbody>`;
    table.innerHTML = html;
    $("#row-count").textContent = `${total.toLocaleString()} shown / ${rows.length.toLocaleString()} rows`;
  }
  function beginLoad() {
    currentFile = fileInput.files[0];
    if (!currentFile) return;
    if (currentFile.size > 100 * 1024 * 1024) { status("Choose a CSV under 100 MB to keep browser memory stable.", true); return; }
    rows = []; headers = []; matches = null; search.value = ""; table.innerHTML = "";
    $("#download").disabled = true;
    parsing = true;
    status(`Reading ${currentFile.name} (${prettySize(currentFile.size)}) in a background worker…`);
    let first = true;
    let stopped = false;
    Papa.parse(currentFile, {
      worker: true, chunkSize: 256 * 1024, skipEmptyLines: "greedy", delimiter: delimiter() || undefined,
      chunk(result, parser) {
        if (result.errors.length) {
          stopped = true; parser.abort(); status(`CSV line ${result.errors[0].row + 1}: ${result.errors[0].message}`, true); return;
        }
        let chunkRows = result.data;
        if (first && chunkRows.length) { headers = uniqueHeaders(chunkRows[0]); chunkRows = chunkRows.slice(1); first = false; }
        for (const row of chunkRows) rows.push(Array.from({ length: headers.length }, (_, index) => row[index] ?? ""));
        if (rows.length > 300000) { stopped = true; parser.abort(); status("Preview limit reached at 300,000 rows. Use a smaller CSV for editing; your original file was not changed.", true); return; }
        render();
        status(`Reading… ${rows.length.toLocaleString()} rows available`);
      },
      complete() {
        parsing = false;
        if (stopped) return;
        if (!headers.length) { status("The file has no rows.", true); return; }
        $("#download").disabled = false;
        render();
        status(`${rows.length.toLocaleString()} rows loaded. Scroll to browse, search to filter, click a cell to edit, then export CSV.`);
      },
      error(error) { parsing = false; status(error.message, true); },
    });
  }
  fileInput.addEventListener("change", beginLoad);
  $("#delimiter").addEventListener("change", () => { if (currentFile) beginLoad(); });
  scroller.addEventListener("scroll", render);
  table.addEventListener("focusout", (event) => {
    const cell = event.target.closest("td[data-col]");
    if (!cell) return;
    rows[Number(cell.parentElement.dataset.row)][Number(cell.dataset.col)] = cell.textContent.replace(/[\r\n]+/g, " ");
  });
  search.addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      const query = search.value.toLocaleLowerCase();
      matches = query ? rows.map((row, index) => row.some((cell) => String(cell).toLocaleLowerCase().includes(query)) ? index : -1).filter((index) => index >= 0) : null;
      scroller.scrollTop = 0; render();
    }, 200);
  });
  $("#download").addEventListener("click", () => {
    if (parsing) return;
    const csv = "\uFEFF" + Papa.unparse([headers, ...rows]);
    download(`${fileName(currentFile)}-edited.csv`, csv, "text/csv;charset=utf-8");
  });
}
function setupChart() {
  let headers = [];
  let rows = [];
  let currentFile;
  let svg = "";
  const colors = ["#1f6feb", "#12a69c", "#855bd6", "#f59e0b", "#ed5b72", "#0ea5e9"];
  const label = (value) => String(value ?? "").slice(0, 24);
  const txt = (value) => escapeHtml(label(value));
  const fail = (message) => {
    svg = "";
    $("#chart-preview").innerHTML = "";
    $("#download-svg").disabled = true;
    $("#download-png").disabled = true;
    status(message, true);
  };
  function options() {
    const selects = [$("#label-column"), $("#value-column")];
    selects.forEach((select) => { select.innerHTML = headers.map((header, index) => `<option value="${index}">${escapeHtml(header)}</option>`).join(""); });
    selects[1].value = String(Math.max(0, headers.findIndex((_, index) => index > 0 && rows.some((row) => Number.isFinite(Number(String(row[index]).replace(/,/g, ""))) && row[index] !== ""))));
  }
  function render() {
    if (!headers.length) return;
    const labelIndex = Number($("#label-column").value);
    const valueIndex = Number($("#value-column").value);
    const groups = new Map();
    for (const row of rows) {
      const value = Number(String(row[valueIndex] ?? "").replace(/,/g, ""));
      if (!Number.isFinite(value) || row[valueIndex] === "") continue;
      const key = String(row[labelIndex] ?? "(blank)") || "(blank)";
      groups.set(key, (groups.get(key) || 0) + value);
    }
    const entries = [...groups].slice(0, 20);
    if (!entries.length) { fail("Choose a numeric value column with at least one valid number."); return; }
    const type = $("#chart-type").value;
    const title = escapeHtml($("#chart-title").value.trim() || `${headers[valueIndex]} by ${headers[labelIndex]}`);
    const width = 960, height = 540;
    svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="${title}"><rect width="960" height="540" fill="white"/><text x="480" y="45" text-anchor="middle" font-family="Arial,sans-serif" font-size="25" font-weight="700" fill="#14243b">${title}</text>`;
    if (type === "pie") {
      if (entries.some(([, value]) => value < 0)) { fail("Pie charts need nonnegative values. Choose bars or a line for this data."); return; }
      const total = entries.reduce((sum, [, value]) => sum + value, 0);
      if (!total) { fail("Pie charts need a positive total."); return; }
      if (entries.filter(([, value]) => value > 0).length === 1) {
        const index = entries.findIndex(([, value]) => value > 0);
        svg += `<circle cx="350" cy="285" r="185" fill="${colors[index % colors.length]}"/>`;
      }
      let angle = -Math.PI / 2;
      entries.forEach(([name, value], index) => {
        if (!value) return;
        const next = angle + value / total * 2 * Math.PI;
        const x1 = 350 + 185 * Math.cos(angle), y1 = 285 + 185 * Math.sin(angle);
        const x2 = 350 + 185 * Math.cos(next), y2 = 285 + 185 * Math.sin(next);
        if (value !== total) svg += `<path d="M350 285 L${x1} ${y1} A185 185 0 ${next - angle > Math.PI ? 1 : 0} 1 ${x2} ${y2} Z" fill="${colors[index % colors.length]}"/>`;
        svg += `<rect x="600" y="${89 + index * 21}" width="12" height="12" fill="${colors[index % colors.length]}"/><text x="620" y="${100 + index * 21}" font-family="Arial,sans-serif" font-size="13" fill="#334155">${txt(name)} (${(value / total * 100).toFixed(1)}%)</text>`;
        angle = next;
      });
    } else {
      const values = entries.map(([, value]) => value);
      const min = Math.min(0, ...values), max = Math.max(0, ...values);
      const span = max - min || 1;
      const y = (value) => 445 - (value - min) / span * 335;
      const baseline = y(0);
      for (let tick = 0; tick <= 4; tick++) {
        const tickValue = min + span * tick / 4;
        const yy = y(tickValue);
        svg += `<line x1="100" y1="${yy}" x2="920" y2="${yy}" stroke="#e2e8f0"/><text x="88" y="${yy + 5}" text-anchor="end" font-family="Arial,sans-serif" font-size="12" fill="#64748b">${Number(tickValue.toPrecision(3)).toLocaleString()}</text>`;
      }
      svg += `<line x1="100" y1="${baseline}" x2="920" y2="${baseline}" stroke="#64748b"/>`;
      const step = 820 / entries.length;
      let points = "";
      entries.forEach(([name, value], index) => {
        const x = 100 + step * (index + .5), yy = y(value);
        if (type === "bar") svg += `<rect x="${x - Math.min(26, step * .34)}" y="${Math.min(yy, baseline)}" width="${Math.min(52, step * .68)}" height="${Math.max(1, Math.abs(baseline - yy))}" fill="${colors[index % colors.length]}"/>`;
        else { points += `${x},${yy} `; svg += `<circle cx="${x}" cy="${yy}" r="5" fill="#1f6feb"/>`; }
        svg += `<text x="${x}" y="468" text-anchor="middle" font-family="Arial,sans-serif" font-size="${entries.length > 12 ? 9 : 11}" fill="#334155">${escapeHtml(label(name).slice(0, entries.length > 10 ? 9 : 15))}</text>`;
      });
      if (type === "line") svg += `<polyline points="${points.trim()}" fill="none" stroke="#1f6feb" stroke-width="3"/>`;
    }
    svg += `</svg>`;
    $("#chart-preview").innerHTML = svg;
    $("#download-svg").disabled = false;
    $("#download-png").disabled = false;
    status(`${entries.length} ${entries.length === 1 ? "category" : "categories"} charted${groups.size > 20 ? `; showing the first 20 of ${groups.size}` : ""}. Values in the same category are summed.`);
  }
  $("#file").addEventListener("change", async (event) => {
    try {
      currentFile = event.target.files[0];
      if (!currentFile) return;
      if (currentFile.name.toLowerCase().endsWith(".xlsx")) {
        const workbook = XLSX.read(await currentFile.arrayBuffer(), { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        ({ headers, rows } = rectangular(XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false })));
      } else ({ headers, rows } = parseCsv(await decodeFile(currentFile, "")));
      if (headers.length < 2 || !rows.length) throw new Error("Use a table with a header row and at least two columns.");
      options(); render();
    } catch (error) { fail(error.message); }
  });
  ["#chart-type", "#label-column", "#value-column"].forEach((selector) => $(selector).addEventListener("change", render));
  $("#chart-title").addEventListener("input", render);
  $("#download-svg").addEventListener("click", () => download(`${fileName(currentFile)}-chart.svg`, svg, "image/svg+xml;charset=utf-8"));
  $("#download-png").addEventListener("click", () => {
    const image = new Image();
    const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
    image.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = 1920; canvas.height = 1080;
      canvas.getContext("2d").drawImage(image, 0, 0, canvas.width, canvas.height);
      URL.revokeObjectURL(url);
      canvas.toBlob((blob) => { if (blob) download(`${fileName(currentFile)}-chart.png`, blob, "image/png"); }, "image/png");
    };
    image.onerror = () => { URL.revokeObjectURL(url); status("PNG export failed in this browser. SVG export is still available.", true); };
    image.src = url;
  });
}
if (page === "json-to-csv") setupJsonToCsv();
if (page === "csv-to-json") setupCsvToJson();
if (page === "csv-to-excel") setupCsvToExcel();
if (page === "csv-viewer") setupViewer();
if (page === "chart-maker") setupChart();
