/**
 * Build an XLSX workbook that contains a real Excel PivotTable.
 *
 * The package includes:
 * - worksheet Source with the loaded rows
 * - worksheet Pivot with one row per distinct row label, in first-seen order,
 *   plus a Grand Total row
 * - xl/pivotCache/pivotCacheDefinition1.xml
 * - xl/pivotCache/pivotCacheRecords1.xml with one record per source data row
 * - xl/pivotTables/pivotTable1.xml named PivotTable1
 *
 * The row field is stored as exact cell text. East and east are different,
 * and spaces stay part of the label. The row field is not trimmed.
 *
 * The value column uses numericText: trim the ends, then accept a plain
 * number. Blank cells, n/a, currency, percents, and thousands separators are
 * left out of the sum. They are not stored as zero. A group with no plain
 * number has a blank sum. The grand total is the sum of the group sums, in
 * group order, using JavaScript numbers. 10.0, 010, and 1e2 are stored as
 * 10, 10, and 100.
 *
 * The workbook does not set refreshOnLoad. The Pivot sheet cells are the
 * sums calculated here. This module does not claim that Excel recalculates
 * them on open.
 */

import { LIMITS, numericText } from './pivot-min-max-distinct.js';

export {
  LIMITS,
  assertByteSize,
  openXlsx,
  parseCsv,
} from './pivot-min-max-distinct.js';

export const PREVIEW_GROUP_LIMIT = 12;
export const SOURCE_SHEET = 'Source';
export const PIVOT_SHEET = 'Pivot';
export const PIVOT_TABLE_NAME = 'PivotTable1';
export const SOURCE_NOTE = 'Browser sums. No refresh-on-load flag.';
export const EXAMPLE_ROW_INDEX = 0;
export const EXAMPLE_VALUE_INDEX = 2;

export const EXAMPLE_CSV = [
  'Region,Product,Amount',
  'East,Desk,10',
  'East,Chair,5',
  'West,Desk,7',
  'West,Lamp,',
  'west,Desk,n/a',
].join('\n');

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const DOS_TIME = 0;
const DOS_DATE = ((2026 - 1980) << 9) | (11 << 5) | 19;
const CELL_CHAR_LIMIT = 32767;
const FIELD_NAME_LIMIT = 255;

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) c = CRC_TABLE[(c ^ bytes[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function concat(chunks) {
  let size = 0;
  for (const chunk of chunks) size += chunk.length;
  const out = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

function u16(value) {
  return Uint8Array.of(value & 255, (value >> 8) & 255);
}

function u32(value) {
  return Uint8Array.of(value & 255, (value >> 8) & 255, (value >> 16) & 255, (value >>> 24) & 255);
}

function zipStore(entries) {
  const encoder = new TextEncoder();
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const entry of entries) {
    const name = encoder.encode(entry.name);
    const data = entry.data;
    const crc = crc32(data);
    const local = concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(0),
      u16(DOS_TIME),
      u16(DOS_DATE),
      u32(crc),
      u32(data.length),
      u32(data.length),
      u16(name.length),
      u16(0),
      name,
      data,
    ]);
    locals.push(local);
    centrals.push(concat([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(0),
      u16(DOS_TIME),
      u16(DOS_DATE),
      u32(crc),
      u32(data.length),
      u32(data.length),
      u16(name.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      name,
    ]));
    offset += local.length;
  }
  const central = concat(centrals);
  const eocd = concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(entries.length),
    u16(entries.length),
    u32(central.length),
    u32(offset),
    u16(0),
  ]);
  return concat([...locals, central, eocd]);
}

function xmlBytes(text) {
  return new TextEncoder().encode(text);
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\t/g, '&#9;')
    .replace(/\n/g, '&#10;')
    .replace(/\r/g, '&#13;');
}

function assertXmlText(value) {
  const text = String(value ?? '');
  if (text.length > CELL_CHAR_LIMIT) {
    throw new Error('A cell is longer than 32,767 characters.');
  }
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        i += 1;
        continue;
      }
      throw new Error('A cell contains a character this XLSX workbook cannot store.');
    }
    if (code >= 0xdc00 && code <= 0xdfff) {
      throw new Error('A cell contains a character this XLSX workbook cannot store.');
    }
    if (code === 0x9 || code === 0xa || code === 0xd) continue;
    if (code < 0x20 || code === 0xfffe || code === 0xffff) {
      throw new Error('A cell contains a character this XLSX workbook cannot store.');
    }
  }
  return text;
}

export function formatPivotNumber(value) {
  if (!Number.isFinite(value)) throw new Error('A pivot sum is not a finite number.');
  const normalized = Object.is(value, -0) ? 0 : value;
  if (normalized === 0) return '0';
  if (Number.isInteger(normalized) && Math.abs(normalized) <= Number.MAX_SAFE_INTEGER) {
    return String(normalized);
  }
  const candidates = [String(normalized), normalized.toPrecision(15), normalized.toPrecision(16)];
  for (const candidate of candidates) {
    const excel = candidate.replace(/e\+/g, 'E').replace(/e-/g, 'E-');
    if (Number(excel) === normalized) return excel;
  }
  throw new Error('A pivot sum could not be stored as an Excel number.');
}

function columnName(index) {
  let n = index + 1;
  let name = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function requireTable(table) {
  if (!table || !Array.isArray(table.headers) || table.headers.length === 0) {
    throw new Error('A header row is required.');
  }
  if (!Array.isArray(table.rows)) throw new Error('Data rows are missing.');
  if (table.headers.length > LIMITS.maxColumns) throw new Error('This file has more than 1,000 columns.');
  if (table.rows.length === 0) throw new Error('The file has a header but no data rows.');
  if (table.rows.length > LIMITS.maxDataRows) throw new Error('This file has more than 100,000 data rows.');
  for (const header of table.headers) {
    assertXmlText(header);
    if (String(header).length > FIELD_NAME_LIMIT) {
      throw new Error('A column name is longer than 255 characters, which an Excel PivotTable field cannot use.');
    }
  }
  return table.headers;
}

function requireIndex(value, width, message) {
  const index = typeof value === 'number' ? value : Number(value);
  if (!Number.isInteger(index) || index < 0 || index >= width) throw new Error(message);
  return index;
}

function cellAt(row, index) {
  if (!Array.isArray(row) || index >= row.length || row[index] == null) return '';
  return String(row[index]);
}

function textShared(rows, index) {
  const items = [];
  const lookup = new Map();
  for (const row of rows) {
    const text = assertXmlText(cellAt(row, index));
    if (!lookup.has(text)) {
      lookup.set(text, items.length);
      items.push(text);
    }
  }
  return { kind: 'text', items, lookup };
}

function valueField(rows, index) {
  const items = [];
  const lookup = new Map();
  let blank = false;
  let integer = true;
  let min = Infinity;
  let max = -Infinity;
  let numberCount = 0;
  for (const row of rows) {
    const raw = assertXmlText(cellAt(row, index));
    if (raw === '') {
      blank = true;
      continue;
    }
    const parsed = numericText(raw);
    if (!parsed) {
      if (!lookup.has(raw)) {
        lookup.set(raw, items.length);
        items.push(raw);
      }
      continue;
    }
    numberCount += 1;
    if (!Number.isInteger(parsed.value)) integer = false;
    if (parsed.value < min) min = parsed.value;
    if (parsed.value > max) max = parsed.value;
  }
  if (numberCount === 0) throw new Error('The value column has no plain number to sum.');
  if (items.length === 0) return { kind: 'number', blank, integer, min, max, items, lookup };
  return { kind: 'mixed', blank, integer, min, max, items, lookup };
}

function recordCell(field, raw) {
  if (field.kind === 'text') return `<x v="${field.lookup.get(raw)}"/>`;
  if (raw === '') return '<m/>';
  const parsed = numericText(raw);
  if (parsed) return `<n v="${formatPivotNumber(parsed.value)}"/>`;
  return `<x v="${field.lookup.get(raw)}"/>`;
}

function sharedItemsXml(field) {
  if (field.kind === 'text') {
    const items = field.items.map((item) => `<s v="${escapeXml(item)}"/>`).join('');
    return `<sharedItems count="${field.items.length}">${items}</sharedItems>`;
  }
  const attrs = [];
  if (field.blank) attrs.push('containsBlank="1"');
  if (field.kind === 'mixed') {
    attrs.push('containsMixedTypes="1"', 'containsSemiMixedTypes="1"', 'containsString="1"', 'containsNumber="1"');
  } else {
    attrs.push('containsSemiMixedTypes="0"', 'containsString="0"', 'containsNumber="1"');
  }
  if (field.integer) attrs.push('containsInteger="1"');
  attrs.push(`minValue="${formatPivotNumber(field.min)}"`, `maxValue="${formatPivotNumber(field.max)}"`);
  if (field.kind === 'number') return `<sharedItems ${attrs.join(' ')}/>`;
  const items = field.items.map((item) => `<s v="${escapeXml(item)}"/>`).join('');
  return `<sharedItems count="${field.items.length}" ${attrs.join(' ')}>${items}</sharedItems>`;
}

function dataCaption(headers, valueIndex) {
  const caption = `Sum of ${headers[valueIndex]}`;
  if (caption.length > FIELD_NAME_LIMIT) {
    throw new Error('The Sum caption for this value column is longer than 255 characters.');
  }
  if (headers.includes(caption)) {
    throw new Error('The Sum caption matches another column name. Rename that column and build the workbook again.');
  }
  return caption;
}

function aggregate(headers, rows, rowIndex, valueIndex) {
  const groups = new Map();
  for (const row of rows) {
    const label = assertXmlText(cellAt(row, rowIndex));
    let group = groups.get(label);
    if (!group) {
      group = { label, sum: 0, numberCount: 0, ignoredCount: 0 };
      groups.set(label, group);
    }
    const raw = assertXmlText(cellAt(row, valueIndex));
    const parsed = raw === '' ? null : numericText(raw);
    if (!parsed) {
      group.ignoredCount += 1;
      continue;
    }
    group.sum = Object.is(group.sum + parsed.value, -0) ? 0 : group.sum + parsed.value;
    group.numberCount += 1;
  }
  const list = [...groups.values()].map((group) => ({
    label: group.label,
    sum: group.numberCount === 0 ? null : group.sum,
    numberCount: group.numberCount,
    ignoredCount: group.ignoredCount,
  }));
  const numberCount = list.reduce((sum, group) => sum + group.numberCount, 0);
  if (numberCount === 0) throw new Error('The value column has no plain number to sum.');
  let grandSum = 0;
  for (const group of list) {
    if (group.sum != null) grandSum = Object.is(grandSum + group.sum, -0) ? 0 : grandSum + group.sum;
  }
  const grandRow = 4 + list.length;
  return {
    groups: list,
    grandSum,
    numberCount,
    ignoredCount: list.reduce((sum, group) => sum + group.ignoredCount, 0),
    recordCount: rows.length,
    rowFieldName: headers[rowIndex],
    valueFieldName: headers[valueIndex],
    dataFieldName: dataCaption(headers, valueIndex),
    location: `A3:B${grandRow}`,
    grandRow,
  };
}

function inlineCell(ref, text) {
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`;
}

function numberCell(ref, value) {
  return `<c r="${ref}"><v>${formatPivotNumber(value)}</v></c>`;
}

function sourceSheetXml(headers, rows, valueIndex) {
  const lastColumn = columnName(headers.length - 1);
  const lastRow = rows.length + 1;
  const lines = [`<row r="1">${headers.map((header, index) => inlineCell(`${columnName(index)}1`, header)).join('')}</row>`];
  rows.forEach((row, rowOffset) => {
    const r = rowOffset + 2;
    const cells = headers.map((_, index) => {
      const raw = cellAt(row, index);
      const ref = `${columnName(index)}${r}`;
      if (index === valueIndex) {
        if (raw === '') return '';
        const parsed = numericText(raw);
        if (parsed) return numberCell(ref, parsed.value);
        return inlineCell(ref, raw);
      }
      if (raw === '') return '';
      return inlineCell(ref, raw);
    }).join('');
    lines.push(`<row r="${r}">${cells}</row>`);
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS_MAIN}"><dimension ref="A1:${lastColumn}${lastRow}"/><sheetData>${lines.join('')}</sheetData></worksheet>`;
}

function pivotSheetXml(model) {
  const lines = [
    `<row r="1">${inlineCell('A1', SOURCE_NOTE)}</row>`,
    `<row r="3">${inlineCell('A3', model.rowFieldName)}${inlineCell('B3', model.dataFieldName)}</row>`,
  ];
  model.groups.forEach((group, index) => {
    const r = 4 + index;
    const sum = group.sum == null ? '' : numberCell(`B${r}`, group.sum);
    lines.push(`<row r="${r}">${inlineCell(`A${r}`, group.label)}${sum}</row>`);
  });
  lines.push(`<row r="${model.grandRow}">${inlineCell(`A${model.grandRow}`, 'Grand Total')}${numberCell(`B${model.grandRow}`, model.grandSum)}</row>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS_MAIN}"><dimension ref="A1:${model.location.split(':')[1]}"/><sheetData>${lines.join('')}</sheetData></worksheet>`;
}

function cacheDefinitionXml(headers, fields, sourceRef, recordCount) {
  const cacheFields = headers.map((header, index) => (
    `<cacheField name="${escapeXml(header)}" numFmtId="0">${sharedItemsXml(fields[index])}</cacheField>`
  )).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_REL}" r:id="rId1" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="${recordCount}" cacheId="1"><cacheSource type="worksheet"><worksheetSource ref="${sourceRef}" sheet="${SOURCE_SHEET}"/></cacheSource><cacheFields count="${headers.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
}

function cacheRecordsXml(rows, fields) {
  const records = rows.map((row) => {
    const cells = fields.map((field, index) => recordCell(field, cellAt(row, index))).join('');
    return `<r>${cells}</r>`;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="${NS_MAIN}" count="${rows.length}">${records}</pivotCacheRecords>`;
}

function pivotTableXml(headers, fields, rowIndex, valueIndex, model) {
  const pivotFields = headers.map((_, index) => {
    if (index === rowIndex) {
      const items = fields[index].items.map((unused, itemIndex) => `<item x="${itemIndex}"/>`).join('');
      return `<pivotField axis="axisRow" showAll="1" sortType="manual" compact="0" outline="0" defaultSubtotal="0"><items count="${fields[index].items.length}">${items}</items></pivotField>`;
    }
    if (index === valueIndex) {
      return '<pivotField dataField="1" showAll="0"><items count="1"><item t="default"/></items></pivotField>';
    }
    return '<pivotField showAll="0"><items count="1"><item t="default"/></items></pivotField>';
  }).join('');
  const rowItems = model.groups.map((unused, index) => `<i><x v="${index}"/></i>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="${NS_MAIN}" name="${PIVOT_TABLE_NAME}" cacheId="1" dataCaption="Values" updatedVersion="3" minRefreshableVersion="3" createdVersion="3" useAutoFormatting="1" itemPrintTitles="1" indent="0" outline="0" outlineData="0" multipleFieldFilters="0" rowGrandTotals="1" colGrandTotals="0" compact="0" compactData="0"><location ref="${model.location}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="0" colPageCount="0"/><pivotFields count="${headers.length}">${pivotFields}</pivotFields><rowFields count="1"><field x="${rowIndex}"/></rowFields><rowItems count="${model.groups.length + 1}">${rowItems}<i t="grand"><x/></i></rowItems><colItems count="1"><i><x/></i></colItems><dataFields count="1"><dataField name="${escapeXml(model.dataFieldName)}" fld="${valueIndex}" subtotal="sum" baseField="0" baseItem="0"/></dataFields></pivotTableDefinition>`;
}

function contentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/><Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}

function workbookXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheets><sheet name="${SOURCE_SHEET}" sheetId="1" r:id="rId1"/><sheet name="${PIVOT_SHEET}" sheetId="2" r:id="rId2"/></sheets><pivotCaches><pivotCache cacheId="1" r:id="rId3"/></pivotCaches></workbook>`;
}

function relationships(pairs) {
  const body = pairs.map(([id, type, target]) => (
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`
  )).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_PKG_REL}">${body}</Relationships>`;
}

export function buildExcelPivotTable(table, rowIndex, valueIndex) {
  const headers = requireTable(table);
  const row = requireIndex(rowIndex, headers.length, 'Choose a row field that exists.');
  const value = requireIndex(valueIndex, headers.length, 'Choose a value column that exists.');
  if (row === value) {
    throw new Error('Choose two different columns. The row field and the value column cannot be the same.');
  }
  for (const dataRow of table.rows) {
    if (!Array.isArray(dataRow)) throw new Error('Data rows are missing.');
  }
  const fields = headers.map((_, index) => (index === value ? valueField(table.rows, index) : textShared(table.rows, index)));
  const model = aggregate(headers, table.rows, row, value);
  const sourceRef = `A1:${columnName(headers.length - 1)}${table.rows.length + 1}`;
  const entries = [
    ['[Content_Types].xml', contentTypesXml()],
    ['_rels/.rels', relationships([
      ['rId1', `${NS_REL}/officeDocument`, 'xl/workbook.xml'],
    ])],
    ['xl/workbook.xml', workbookXml()],
    ['xl/_rels/workbook.xml.rels', relationships([
      ['rId1', `${NS_REL}/worksheet`, 'worksheets/sheet1.xml'],
      ['rId2', `${NS_REL}/worksheet`, 'worksheets/sheet2.xml'],
      ['rId3', `${NS_REL}/pivotCacheDefinition`, 'pivotCache/pivotCacheDefinition1.xml'],
      ['rId4', `${NS_REL}/styles`, 'styles.xml'],
    ])],
    ['xl/styles.xml', stylesXml()],
    ['xl/worksheets/sheet1.xml', sourceSheetXml(headers, table.rows, value)],
    ['xl/worksheets/sheet2.xml', pivotSheetXml(model)],
    ['xl/worksheets/_rels/sheet2.xml.rels', relationships([
      ['rId1', `${NS_REL}/pivotTable`, '../pivotTables/pivotTable1.xml'],
    ])],
    ['xl/pivotCache/pivotCacheDefinition1.xml', cacheDefinitionXml(headers, fields, sourceRef, table.rows.length)],
    ['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', relationships([
      ['rId1', `${NS_REL}/pivotCacheRecords`, 'pivotCacheRecords1.xml'],
    ])],
    ['xl/pivotCache/pivotCacheRecords1.xml', cacheRecordsXml(table.rows, fields)],
    ['xl/pivotTables/pivotTable1.xml', pivotTableXml(headers, fields, row, value, model)],
    ['xl/pivotTables/_rels/pivotTable1.xml.rels', relationships([
      ['rId1', `${NS_REL}/pivotCacheDefinition`, '../pivotCache/pivotCacheDefinition1.xml'],
    ])],
  ].map(([name, text]) => ({ name, data: xmlBytes(text) }));
  return {
    bytes: zipStore(entries),
    fileName: 'excel-pivot-table.xlsx',
    sourceRef,
    ...model,
  };
}

export function previewGroups(result, limit = PREVIEW_GROUP_LIMIT) {
  return {
    groups: result.groups.slice(0, limit),
    hiddenGroupCount: Math.max(0, result.groups.length - limit),
  };
}
