/**
 * Grouped minimum, maximum, and distinct count.
 *
 * Distinct count uses the exact cell text. Empty cells share one value, the
 * empty string. Spaces and letter case are kept. A completely blank record is
 * skipped on import.
 *
 * Minimum and maximum use a cell only when its trimmed text is a finite plain
 * number (optional sign, decimal, or exponent). Thousands separators, currency,
 * percents, words, and blanks are ignored. The displayed extreme is the trimmed
 * text of the first row that reaches that number. A group with no such number
 * gets blank minimum and maximum cells, not zero.
 *
 * The last result row is the whole file. Its distinct count is not the sum of
 * the group counts. Files are read locally by the caller; this module does not
 * upload them.
 */

export const LIMITS = {
  maxBytes: 10 * 1024 * 1024,
  maxDataRows: 100000,
  maxColumns: 1000,
};

export const RESULT_HEADERS = ['Group', 'Minimum', 'Maximum', 'Distinct count', 'Scope'];
export const OVERALL_GROUP_LABEL = '(all rows)';
export const SCOPE_GROUP = 'group';
export const SCOPE_OVERALL = 'all rows';
export const PREVIEW_GROUP_LIMIT = 12;

export const EXAMPLE_CSV = [
  'Region,Amount',
  'East,5',
  'East,1',
  'East,5',
  'West,5',
  'West,9',
  'West,n/a',
  'North,',
  'North,2',
].join('\n');

const NUMERIC = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;
const MAX_ENTRY_BYTES = 32 * 1024 * 1024;

export function assertByteSize(byteLength) {
  if (!Number.isFinite(byteLength) || byteLength < 0) {
    throw new Error('This file could not be read.');
  }
  if (byteLength > LIMITS.maxBytes) {
    throw new Error('This file is larger than 10 MB.');
  }
}

export function normalizeHeaders(cells) {
  const used = new Set();
  return cells.map((cell, index) => {
    const raw = cell == null ? '' : String(cell);
    const base = raw === '' ? `Column ${index + 1}` : raw;
    let name = base;
    let n = 2;
    while (used.has(name)) {
      name = `${base} (${n})`;
      n += 1;
    }
    used.add(name);
    return name;
  });
}

export function numericText(raw) {
  const trimmed = String(raw ?? '').trim();
  if (!NUMERIC.test(trimmed)) return null;
  const value = Number(trimmed);
  if (!Number.isFinite(value)) return null;
  return { value, text: trimmed };
}

function cellString(value) {
  return value == null ? '' : String(value);
}

function matrixToTable(matrix) {
  if (!matrix.length) throw new Error('The selected worksheet is empty.');
  const width = matrix.reduce((max, row) => Math.max(max, row.length), 0);
  if (width > LIMITS.maxColumns) {
    throw new Error('This file has more than 1,000 columns.');
  }
  const padded = matrix.map((row) => {
    const copy = row.slice(0, width);
    while (copy.length < width) copy.push('');
    return copy;
  });
  const headers = normalizeHeaders(padded[0]);
  const rows = padded.slice(1).filter((row) => row.some((cell) => cell !== ''));
  if (rows.length > LIMITS.maxDataRows) {
    throw new Error('This file has more than 100,000 data rows.');
  }
  if (rows.length === 0) {
    throw new Error('The file has a header but no data rows.');
  }
  return { headers, rows };
}

export function detectDelimiter(text) {
  const line = String(text).split(/\r\n|\n|\r/).find((item) => item.trim() !== '') ?? '';
  const candidates = [',', ';', '\t', '|'];
  let best = ',';
  let bestCount = -1;
  for (const candidate of candidates) {
    const count = countDelimiter(line, candidate);
    if (count > bestCount) {
      best = candidate;
      bestCount = count;
    }
  }
  return best;
}

function countDelimiter(line, delimiter) {
  let count = 0;
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        i += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (!inQuotes && ch === delimiter) count += 1;
  }
  return count;
}

export function parseCsv(text, delimiter) {
  if (text == null) throw new Error('The file is empty.');
  let input = String(text);
  if (input.charCodeAt(0) === 0xfeff) input = input.slice(1);
  if (input.trim() === '') throw new Error('The file is empty.');

  const delim = !delimiter || delimiter === 'auto' ? detectDelimiter(input) : delimiter;
  if (![',', ';', '\t', '|'].includes(delim)) {
    throw new Error('Choose a comma, semicolon, tab, or pipe delimiter.');
  }

  const records = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];
    if (inQuotes) {
      if (ch === '"') {
        if (input[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"' && field === '') {
      inQuotes = true;
      continue;
    }
    if (ch === delim) {
      row.push(field);
      field = '';
      continue;
    }
    if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && input[i + 1] === '\n') i += 1;
      row.push(field);
      field = '';
      records.push(row);
      row = [];
      continue;
    }
    field += ch;
  }
  if (inQuotes) {
    throw new Error('This CSV could not be parsed because a quote is not closed.');
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    records.push(row);
  }

  const kept = records.filter((record) => record.some((cell) => cell !== ''));
  if (kept.length === 0) throw new Error('The file is empty.');
  return { ...matrixToTable(kept), delimiter: delim };
}

function csvEscape(value) {
  const text = cellString(value);
  if (/[",\r\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

export function serializeCsv(headers, rows) {
  const lines = [headers, ...rows].map((row) => row.map(csvEscape).join(','));
  return `\uFEFF${lines.join('\n')}\n`;
}

function consider(current, candidate, direction) {
  if (!candidate) return current;
  if (!current) return candidate;
  if (direction === 'min' && candidate.value < current.value) return candidate;
  if (direction === 'max' && candidate.value > current.value) return candidate;
  return current;
}

export function aggregateMinMaxDistinct(table, groupIndex, valueIndex) {
  if (!table || !Array.isArray(table.headers) || table.headers.length === 0) {
    throw new Error('A header row is required.');
  }
  if (!Array.isArray(table.rows)) throw new Error('Data rows are missing.');
  const width = table.headers.length;
  if (!Number.isInteger(groupIndex) || groupIndex < 0 || groupIndex >= width) {
    throw new Error('Choose a row field from the header.');
  }
  if (!Number.isInteger(valueIndex) || valueIndex < 0 || valueIndex >= width) {
    throw new Error('Choose a value field from the header.');
  }
  if (table.rows.length === 0) {
    throw new Error('The file has a header but no data rows.');
  }

  const groups = new Map();
  const overallDistinct = new Set();
  let overallMin = null;
  let overallMax = null;

  for (const row of table.rows) {
    if (!Array.isArray(row)) throw new Error('A data row is not a list of cells.');
    const group = cellString(row[groupIndex]);
    const value = cellString(row[valueIndex]);
    let bucket = groups.get(group);
    if (!bucket) {
      bucket = { distinct: new Set(), min: null, max: null };
      groups.set(group, bucket);
    }
    bucket.distinct.add(value);
    overallDistinct.add(value);
    const numeric = numericText(value);
    bucket.min = consider(bucket.min, numeric, 'min');
    bucket.max = consider(bucket.max, numeric, 'max');
    overallMin = consider(overallMin, numeric, 'min');
    overallMax = consider(overallMax, numeric, 'max');
  }

  const rows = [];
  for (const [group, bucket] of groups) {
    rows.push([
      group,
      bucket.min ? bucket.min.text : '',
      bucket.max ? bucket.max.text : '',
      String(bucket.distinct.size),
      SCOPE_GROUP,
    ]);
  }
  rows.push([
    OVERALL_GROUP_LABEL,
    overallMin ? overallMin.text : '',
    overallMax ? overallMax.text : '',
    String(overallDistinct.size),
    SCOPE_OVERALL,
  ]);

  return {
    headers: RESULT_HEADERS.slice(),
    rows,
    groupCount: groups.size,
    dataRowCount: table.rows.length,
  };
}

export function previewResult(result, limit = PREVIEW_GROUP_LIMIT) {
  const groups = result.rows.filter((row) => row[4] === SCOPE_GROUP);
  const overall = result.rows[result.rows.length - 1];
  if (groups.length <= limit) return result.rows.slice();
  return [...groups.slice(0, limit), overall];
}

function isNodeRuntime() {
  return typeof process !== 'undefined' && process.versions != null && process.versions.node != null;
}

async function inflateRaw(data) {
  if (isNodeRuntime()) {
    const { inflateRawSync } = await import('node:zlib');
    return new Uint8Array(inflateRawSync(data));
  }
  if (typeof DecompressionStream !== 'function') {
    throw new Error('This browser cannot read XLSX files. Save the sheet as CSV, or use a current version of Chrome, Edge, or Firefox.');
  }
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function toUint8(bytes) {
  if (bytes instanceof Uint8Array) return bytes;
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  throw new Error('This XLSX file could not be read.');
}

function u16(u8, offset) {
  return u8[offset] | (u8[offset + 1] << 8);
}

function u32(u8, offset) {
  return ((u8[offset] | (u8[offset + 1] << 8) | (u8[offset + 2] << 16) | (u8[offset + 3] << 24)) >>> 0);
}

function readCentralDirectory(u8) {
  const min = Math.max(0, u8.length - (22 + 65535));
  let eocd = -1;
  for (let i = u8.length - 22; i >= min; i -= 1) {
    if (u8[i] !== 0x50 || u8[i + 1] !== 0x4b || u8[i + 2] !== 0x05 || u8[i + 3] !== 0x06) continue;
    const commentLength = u16(u8, i + 20);
    if (i + 22 + commentLength !== u8.length) continue;
    eocd = i;
    break;
  }
  if (eocd < 0) throw new Error('This XLSX file could not be read.');

  const count = u16(u8, eocd + 10);
  let cursor = u32(u8, eocd + 16);
  const entries = new Map();
  for (let n = 0; n < count; n += 1) {
    if (cursor + 46 > u8.length || u32(u8, cursor) !== 0x02014b50) {
      throw new Error('This XLSX file could not be read.');
    }
    const method = u16(u8, cursor + 10);
    const compSize = u32(u8, cursor + 20);
    const uncompSize = u32(u8, cursor + 24);
    const nameLength = u16(u8, cursor + 28);
    const extraLength = u16(u8, cursor + 30);
    const commentLength = u16(u8, cursor + 32);
    const localOffset = u32(u8, cursor + 42);
    const nameStart = cursor + 46;
    const nameEnd = nameStart + nameLength;
    if (nameEnd > u8.length) throw new Error('This XLSX file could not be read.');
    const name = new TextDecoder('utf-8').decode(u8.subarray(nameStart, nameEnd)).replace(/\\/g, '/');
    if (uncompSize > MAX_ENTRY_BYTES || compSize > MAX_ENTRY_BYTES) {
      throw new Error('This XLSX file is too large to read in the browser.');
    }
    entries.set(name, { method, compSize, uncompSize, localOffset });
    cursor = nameEnd + extraLength + commentLength;
  }
  return entries;
}

async function readZipText(u8, entries, path) {
  const entry = entries.get(path);
  if (!entry) return null;
  if (entry.localOffset + 30 > u8.length || u32(u8, entry.localOffset) !== 0x04034b50) {
    throw new Error('This XLSX file could not be read.');
  }
  const nameLength = u16(u8, entry.localOffset + 26);
  const extraLength = u16(u8, entry.localOffset + 28);
  const start = entry.localOffset + 30 + nameLength + extraLength;
  const end = start + entry.compSize;
  if (end > u8.length) throw new Error('This XLSX file could not be read.');
  const compressed = u8.subarray(start, end);
  let raw;
  if (entry.method === 0) raw = compressed;
  else if (entry.method === 8) raw = await inflateRaw(compressed);
  else throw new Error('This XLSX file could not be read.');
  if (raw.length > MAX_ENTRY_BYTES) {
    throw new Error('This XLSX file is too large to read in the browser.');
  }
  return new TextDecoder('utf-8').decode(raw).replace(/^\uFEFF/, '');
}

function decodeXml(text) {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => {
      const code = Number.parseInt(hex, 16);
      if (!Number.isInteger(code) || code < 0 || code > 0x10ffff) {
        throw new Error('This XLSX file could not be read.');
      }
      return String.fromCodePoint(code);
    })
    .replace(/&#(\d+);/g, (_, num) => {
      const code = Number(num);
      if (!Number.isInteger(code) || code < 0 || code > 0x10ffff) {
        throw new Error('This XLSX file could not be read.');
      }
      return String.fromCodePoint(code);
    })
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function attr(tag, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = new RegExp(`\\s${escaped}\\s*=\\s*("([^"]*)"|'([^']*)')`).exec(` ${tag}`);
  if (!match) return '';
  return decodeXml(match[2] ?? match[3] ?? '');
}

function extractBlocks(xml, localName) {
  const blocks = [];
  const startRe = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b`, 'g');
  let match = startRe.exec(xml);
  while (match) {
    const start = match.index;
    const gt = xml.indexOf('>', start);
    if (gt < 0) break;
    const tag = xml.slice(start, gt + 1);
    if (tag.endsWith('/>')) {
      blocks.push({ attrs: tag, inner: '' });
      const next = gt + 1;
      startRe.lastIndex = next > start ? next : start + 1;
    } else {
      const closeRe = new RegExp(`</(?:[A-Za-z_][\\w.-]*:)?${localName}>`, 'g');
      closeRe.lastIndex = gt + 1;
      const close = closeRe.exec(xml);
      if (!close) break;
      blocks.push({ attrs: tag, inner: xml.slice(gt + 1, close.index) });
      const next = close.index + close[0].length;
      startRe.lastIndex = next > start ? next : start + 1;
    }
    match = startRe.exec(xml);
  }
  return blocks;
}

function collectText(xml) {
  const cleaned = xml.replace(/<rPh\b[^>]*>[\s\S]*?<\/rPh>/g, '');
  return extractBlocks(cleaned, 't').map((part) => decodeXml(part.inner)).join('');
}

export function sharedStringsFromXml(xml) {
  if (!xml) return [];
  const cleaned = xml.replace(/<rPh\b[^>]*>[\s\S]*?<\/rPh>/g, '');
  return extractBlocks(cleaned, 'si').map((item) => collectText(item.inner));
}

export function columnIndex(ref) {
  const match = /^([A-Z]+)/.exec(String(ref));
  if (!match) throw new Error('This XLSX file could not be read.');
  let n = 0;
  for (const ch of match[1]) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
    if (n > 16384) throw new Error('This XLSX file could not be read.');
  }
  return n - 1;
}

function cellText(cell, sharedStrings) {
  const type = attr(cell.attrs, 't');
  if (type === 'inlineStr') {
    const inline = extractBlocks(cell.inner, 'is');
    const source = inline.length ? inline.map((part) => part.inner).join('') : cell.inner;
    return collectText(source);
  }
  const values = extractBlocks(cell.inner, 'v');
  const raw = values.length ? decodeXml(values.map((part) => part.inner).join('')) : '';
  if (type === 's') {
    const index = Number.parseInt(raw.trim(), 10);
    if (!Number.isInteger(index) || index < 0 || index >= sharedStrings.length) {
      throw new Error('This XLSX file has a text reference this page could not resolve.');
    }
    return sharedStrings[index];
  }
  if (type === 'b') {
    if (raw.trim() === '1') return 'TRUE';
    if (raw.trim() === '0') return 'FALSE';
    return raw;
  }
  return raw;
}

export function parseWorksheetXml(xml, sharedStrings = []) {
  const rows = [];
  for (const rowBlock of extractBlocks(xml, 'row')) {
    const cells = extractBlocks(rowBlock.inner, 'c');
    if (cells.length === 0) continue;
    const values = [];
    let next = 0;
    for (const cell of cells) {
      const ref = attr(cell.attrs, 'r');
      const col = ref ? columnIndex(ref) : next;
      if (col < 0) throw new Error('This XLSX file could not be read.');
      const text = cellText(cell, sharedStrings);
      while (values.length < col) values.push('');
      values[col] = text;
      next = col + 1;
    }
    if (values.some((value) => value !== '')) rows.push(values);
  }
  return rows;
}

function resolveSheetPath(target) {
  let path = target.replace(/\\/g, '/');
  try {
    path = decodeURIComponent(path);
  } catch {
    throw new Error('This XLSX file could not be read.');
  }
  if (path.startsWith('/')) path = path.slice(1);
  const parts = [];
  for (const part of path.split('/')) {
    if (part === '' || part === '.') continue;
    if (part === '..') parts.pop();
    else parts.push(part);
  }
  path = parts.join('/');
  if (!path.startsWith('xl/')) path = `xl/${path}`;
  return path;
}

export async function openXlsx(bytes) {
  const u8 = toUint8(bytes);
  assertByteSize(u8.byteLength);
  if (u8.length < 4 || u8[0] !== 0x50 || u8[1] !== 0x4b) {
    throw new Error('This XLSX file could not be read.');
  }
  const entries = readCentralDirectory(u8);
  const workbookXml = await readZipText(u8, entries, 'xl/workbook.xml');
  const relsXml = await readZipText(u8, entries, 'xl/_rels/workbook.xml.rels');
  if (!workbookXml || !relsXml) throw new Error('This XLSX file could not be read.');

  const rels = new Map();
  for (const rel of extractBlocks(relsXml, 'Relationship')) {
    const id = attr(rel.attrs, 'Id');
    const target = attr(rel.attrs, 'Target');
    if (id && target) rels.set(id, resolveSheetPath(target));
  }

  const sheets = extractBlocks(workbookXml, 'sheet').map((sheet, index) => {
    const name = attr(sheet.attrs, 'name') || `Sheet ${index + 1}`;
    const id = attr(sheet.attrs, 'r:id');
    const path = rels.get(id);
    if (!path) throw new Error('This XLSX file could not be read.');
    return { name, path };
  });
  if (sheets.length === 0) throw new Error('This XLSX file could not be read.');

  let shared = null;
  async function loadSharedStrings() {
    if (shared) return shared;
    const xml = await readZipText(u8, entries, 'xl/sharedStrings.xml');
    shared = xml ? sharedStringsFromXml(xml) : [];
    return shared;
  }

  const cache = new Map();
  return {
    sheetNames: sheets.map((sheet) => sheet.name),
    async sheet(index) {
      if (!Number.isInteger(index) || index < 0 || index >= sheets.length) {
        throw new Error('Choose a worksheet that exists in the workbook.');
      }
      if (cache.has(index)) return cache.get(index);
      const xml = await readZipText(u8, entries, sheets[index].path);
      if (!xml) throw new Error('This XLSX file could not be read.');
      const table = matrixToTable(parseWorksheetXml(xml, await loadSharedStrings()));
      const result = { name: sheets[index].name, ...table };
      cache.set(index, result);
      return result;
    },
  };
}
