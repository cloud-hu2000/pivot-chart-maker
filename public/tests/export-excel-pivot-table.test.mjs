import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import {
  EXAMPLE_CSV,
  EXAMPLE_ROW_INDEX,
  EXAMPLE_VALUE_INDEX,
  PIVOT_TABLE_NAME,
  SOURCE_NOTE,
  buildExcelPivotTable,
  formatPivotNumber,
  openXlsx,
  parseCsv,
} from '../js/export-excel-pivot-table.js';

function zipEntries(bytes) {
  const dir = mkdtempSync(join(tmpdir(), 'pivot-xlsx-'));
  const file = join(dir, 'book.xlsx');
  writeFileSync(file, bytes);
  try {
    const raw = execFileSync('python3', ['-c', `
import json, sys, zipfile
with zipfile.ZipFile(sys.argv[1]) as archive:
    json.dump({name: archive.read(name).decode("utf-8") for name in archive.namelist()}, sys.stdout)
`, file], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 });
    return new Map(Object.entries(JSON.parse(raw)));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

function text(files, name) {
  const value = files.get(name);
  assert.ok(value, `missing ${name}`);
  return value;
}

function tagCount(xml, tag) {
  return [...xml.matchAll(new RegExp(`<${tag}\\b`, 'g'))].length;
}

const table = parseCsv(EXAMPLE_CSV, ',');
const built = buildExcelPivotTable(table, EXAMPLE_ROW_INDEX, EXAMPLE_VALUE_INDEX);
assert.equal(built.groups.length, 3);
assert.equal(built.groups[0].label, 'East');
assert.equal(built.groups[0].sum, 15);
assert.equal(built.groups[0].numberCount, 2);
assert.equal(built.groups[0].ignoredCount, 0);
assert.equal(built.groups[1].label, 'West');
assert.equal(built.groups[1].sum, 7);
assert.equal(built.groups[1].numberCount, 1);
assert.equal(built.groups[1].ignoredCount, 1);
assert.equal(built.groups[2].label, 'west');
assert.equal(built.groups[2].sum, null);
assert.equal(built.groups[2].numberCount, 0);
assert.equal(built.groups[2].ignoredCount, 1);
assert.equal(built.grandSum, 22);
assert.equal(built.numberCount, 3);
assert.equal(built.ignoredCount, 2);
assert.equal(built.recordCount, 5);
assert.equal(built.location, 'A3:B7');
assert.equal(built.dataFieldName, 'Sum of Amount');
assert.equal(built.bytes[0], 0x50);
assert.equal(built.bytes[1], 0x4b);

const files = zipEntries(built.bytes);
const required = [
  '[Content_Types].xml',
  '_rels/.rels',
  'xl/workbook.xml',
  'xl/_rels/workbook.xml.rels',
  'xl/styles.xml',
  'xl/worksheets/sheet1.xml',
  'xl/worksheets/sheet2.xml',
  'xl/worksheets/_rels/sheet2.xml.rels',
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
  'xl/pivotCache/pivotCacheRecords1.xml',
  'xl/pivotTables/pivotTable1.xml',
  'xl/pivotTables/_rels/pivotTable1.xml.rels',
];
assert.deepEqual([...files.keys()].sort(), [...required].sort());

const contentTypes = text(files, '[Content_Types].xml');
assert.match(contentTypes, /pivotCacheDefinition\+xml/);
assert.match(contentTypes, /pivotCacheRecords\+xml/);
assert.match(contentTypes, /spreadsheetml\.pivotTable\+xml/);
assert.match(contentTypes, /PartName="\/xl\/pivotCache\/pivotCacheDefinition1\.xml"/);
assert.match(contentTypes, /PartName="\/xl\/pivotCache\/pivotCacheRecords1\.xml"/);
assert.match(contentTypes, /PartName="\/xl\/pivotTables\/pivotTable1\.xml"/);

const workbook = text(files, 'xl/workbook.xml');
assert.match(workbook, /<sheet name="Source"/);
assert.match(workbook, /<sheet name="Pivot"/);
assert.match(workbook, /<pivotCache cacheId="1"/);
assert.match(text(files, 'xl/_rels/workbook.xml.rels'), /relationships\/pivotCacheDefinition" Target="pivotCache\/pivotCacheDefinition1\.xml"/);
assert.match(text(files, 'xl/worksheets/_rels/sheet2.xml.rels'), /relationships\/pivotTable" Target="\.\.\/pivotTables\/pivotTable1\.xml"/);
assert.match(text(files, 'xl/pivotTables/_rels/pivotTable1.xml.rels'), /relationships\/pivotCacheDefinition" Target="\.\.\/pivotCache\/pivotCacheDefinition1\.xml"/);
assert.match(text(files, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'), /relationships\/pivotCacheRecords" Target="pivotCacheRecords1\.xml"/);

const cache = text(files, 'xl/pivotCache/pivotCacheDefinition1.xml');
assert.equal(/refreshOnLoad\s*=/.test(cache), false);
assert.equal(/saveData\s*=\s*"0"/.test(cache), false);
assert.match(cache, /recordCount="5"/);
assert.match(cache, /cacheId="1"/);
assert.match(cache, /worksheetSource ref="A1:C6" sheet="Source"/);
assert.match(cache, /<cacheField name="Region"[\s\S]*?<s v="East"\/><s v="West"\/><s v="west"\//);
assert.match(cache, /<cacheField name="Product"[\s\S]*?<s v="Desk"\/><s v="Chair"\/><s v="Lamp"\//);
assert.match(cache, /<cacheField name="Amount"[\s\S]*?containsString="1"[\s\S]*?containsNumber="1"[\s\S]*?<s v="n\/a"\//);
assert.match(cache, /minValue="5"/);
assert.match(cache, /maxValue="10"/);
assert.match(cache, /containsBlank="1"/);
assert.match(cache, /containsInteger="1"/);

const records = text(files, 'xl/pivotCache/pivotCacheRecords1.xml');
assert.equal(tagCount(records, 'r'), 5);
assert.match(records, /count="5"/);
assert.equal(tagCount(records, 'm'), 1);
assert.match(records, /<n v="10"\/>/);
assert.match(records, /<n v="5"\/>/);
assert.match(records, /<n v="7"\/>/);
assert.equal(/refreshOnLoad\s*=/.test(records), false);

const pivot = text(files, 'xl/pivotTables/pivotTable1.xml');
assert.match(pivot, new RegExp(`name="${PIVOT_TABLE_NAME}"`));
assert.match(pivot, /cacheId="1"/);
assert.match(pivot, /location ref="A3:B7"/);
assert.match(pivot, /<rowFields count="1"><field x="0"\/>/);
assert.match(pivot, /<dataField name="Sum of Amount" fld="2" subtotal="sum"/);
assert.match(pivot, /<i t="grand">/);
assert.equal(/refreshOnLoad\s*=/.test(pivot), false);
assert.equal(tagCount(pivot, 'pivotField'), 3);

const pivotSheet = text(files, 'xl/worksheets/sheet2.xml');
assert.match(pivotSheet, new RegExp(SOURCE_NOTE.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
assert.match(pivotSheet, /<c r="A4"[\s\S]*?>East</);
assert.match(pivotSheet, /<c r="B4"><v>15<\/v><\/c>/);
assert.match(pivotSheet, /<c r="B5"><v>7<\/v><\/c>/);
assert.equal(/<c r="B6"/.test(pivotSheet), false);
assert.match(pivotSheet, /<c r="A6"[\s\S]*?>west</);
assert.match(pivotSheet, /<c r="A7"[\s\S]*?>Grand Total</);
assert.match(pivotSheet, /<c r="B7"><v>22<\/v><\/c>/);

for (const xml of files.values()) {
  assert.equal(/refreshOnLoad\s*=/.test(xml), false);
}

const opened = await openXlsx(built.bytes);
assert.deepEqual(opened.sheetNames, ['Source', 'Pivot']);
const source = await opened.sheet(0);
assert.deepEqual(source.headers, ['Region', 'Product', 'Amount']);
assert.deepEqual(source.rows, [
  ['East', 'Desk', '10'],
  ['East', 'Chair', '5'],
  ['West', 'Desk', '7'],
  ['West', 'Lamp', ''],
  ['west', 'Desk', 'n/a'],
]);

const byProduct = buildExcelPivotTable(table, 1, 2);
assert.deepEqual(byProduct.groups.map((group) => [group.label, group.sum]), [
  ['Desk', 17],
  ['Chair', 5],
  ['Lamp', null],
]);
assert.equal(byProduct.grandSum, 22);
const productPivot = text(zipEntries(byProduct.bytes), 'xl/pivotTables/pivotTable1.xml');
assert.match(productPivot, /<rowFields count="1"><field x="1"\/>/);
assert.match(productPivot, /fld="2" subtotal="sum"/);

const zeros = buildExcelPivotTable(parseCsv('Region,Amount\nEast,1\nEast,-1\n', ','), 0, 1);
assert.equal(zeros.groups[0].sum, 0);
assert.equal(zeros.grandSum, 0);
const zeroSheet = text(zipEntries(zeros.bytes), 'xl/worksheets/sheet2.xml');
assert.match(zeroSheet, /<c r="B4"><v>0<\/v><\/c>/);
assert.match(zeroSheet, /<c r="B5"><v>0<\/v><\/c>/);

const pure = buildExcelPivotTable(parseCsv('Region,Amount\nEast,1\nEast,2\nWest,3\n', ','), 0, 1);
assert.equal(pure.grandSum, 6);
const pureCache = text(zipEntries(pure.bytes), 'xl/pivotCache/pivotCacheDefinition1.xml');
assert.match(pureCache, /<cacheField name="Amount"[^>]*>\s*<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="3"\/>/);
assert.equal(/<cacheField name="Amount"[\s\S]*?<s /.test(pureCache), false);

const decimal = buildExcelPivotTable(parseCsv('Region,Amount\nEast,1.5\nEast,2\n', ','), 0, 1);
const decimalCache = text(zipEntries(decimal.bytes), 'xl/pivotCache/pivotCacheDefinition1.xml');
assert.match(decimalCache, /containsNumber="1"/);
assert.equal(/containsInteger="1"/.test(decimalCache), false);
assert.equal(decimal.grandSum, 3.5);

const binary = buildExcelPivotTable(parseCsv('Region,Amount\nEast,0.1\nEast,0.2\n', ','), 0, 1);
assert.equal(binary.grandSum, 0.1 + 0.2);
assert.equal(Number(formatPivotNumber(binary.grandSum)), binary.grandSum);
const binarySheet = text(zipEntries(binary.bytes), 'xl/worksheets/sheet2.xml');
assert.match(binarySheet, new RegExp(`<v>${formatPivotNumber(binary.grandSum).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}<\\/v>`));

const spaced = buildExcelPivotTable(parseCsv('Region,Amount\nEast,10\nEast , 8 \nwest,4\n', ','), 0, 1);
assert.deepEqual(spaced.groups.map((group) => [group.label, group.sum]), [
  ['East', 10],
  ['East ', 8],
  ['west', 4],
]);

const codes = buildExcelPivotTable(parseCsv('Code,Amount\n01,010\n01,1e2\n02,10.0\n', ','), 0, 1);
assert.deepEqual(codes.groups.map((group) => group.label), ['01', '02']);
assert.equal(codes.groups[0].sum, 110);
assert.equal(codes.groups[1].sum, 10);
const codeSource = await openXlsx(codes.bytes).then((book) => book.sheet(0));
assert.deepEqual(codeSource.rows.map((row) => row[0]), ['01', '01', '02']);
assert.deepEqual(codeSource.rows.map((row) => row[1]), ['10', '100', '10']);

const ignored = buildExcelPivotTable(parseCsv('Region,Amount\nEast,"1,000"\nEast,$5\nEast,5%\nEast,2\n', ','), 0, 1);
assert.equal(ignored.groups[0].sum, 2);
assert.equal(ignored.groups[0].numberCount, 1);
assert.equal(ignored.groups[0].ignoredCount, 3);

const escaped = buildExcelPivotTable(parseCsv('A&B,Amount\nM&M,4\nM&M,1\n', ','), 0, 1);
const escapedCache = text(zipEntries(escaped.bytes), 'xl/pivotCache/pivotCacheDefinition1.xml');
assert.match(escapedCache, /cacheField name="A&amp;B"/);
assert.match(escapedCache, /<s v="M&amp;M"\/>/);
assert.equal(escaped.grandSum, 5);
assert.equal(escaped.dataFieldName, 'Sum of Amount');

assert.throws(() => buildExcelPivotTable(table, 0, 0), /two different columns/);
assert.throws(() => buildExcelPivotTable(table, 2, 2), /two different columns/);
assert.throws(() => buildExcelPivotTable(table, 5, 2), /row field that exists/);
assert.throws(() => buildExcelPivotTable(table, 0, 9), /value column that exists/);
assert.throws(() => buildExcelPivotTable({ headers: ['Region', 'Amount'], rows: [['East', 'n/a'], ['West', '']] }, 0, 1), /no plain number/);
assert.throws(() => buildExcelPivotTable({ headers: ['Region'], rows: [['East']] }, 0, 0), /two different columns/);
assert.throws(() => buildExcelPivotTable({ headers: [], rows: [] }, 0, 1), /header row is required/);
assert.throws(() => buildExcelPivotTable({ headers: ['Region', 'Amount'], rows: [] }, 0, 1), /no data rows/);
assert.throws(() => buildExcelPivotTable({ headers: ['Region', 'Amount'], rows: [['East', '1\u0001']] }, 0, 1), /cannot store/);
assert.throws(() => buildExcelPivotTable({
  headers: ['Amount', 'Sum of Amount'],
  rows: [['2', 'x']],
}, 1, 0), /Sum caption matches another column/);
assert.throws(() => parseCsv('Region,Amount\n"East,1\n', ','), /quote is not closed/);

const html = readFileSync(new URL('../export-excel-pivot-table.html', import.meta.url), 'utf8');
const title = html.match(/<title>([^<]*)<\/title>/)[1];
const description = html.match(/<meta name="description" content="([^"]*)"/)[1];
const lede = html.match(/<p class="lede">([^<]*)<\/p>/)[1];
assert.equal(title, 'Export Excel PivotTable XLSX');
assert.equal([...html.matchAll(/<h1>/g)].length, 1);
assert.equal(html.includes('<h1>Export Excel PivotTable XLSX</h1>'), true);
assert.equal(title.length + ' | Pivot Chart Maker'.length <= 60, true);
assert.equal(description.length >= 140 && description.length <= 160, true, String(description.length));
assert.notEqual(description, lede);
assert.match(html, /https:\/\/www\.pivotchartmaker\.org\/export-excel-pivot-table/);
assert.match(html, /pivot cache records/);
assert.match(html, /PivotTable1/);
assert.match(html, /Browser sums\. No refresh-on-load flag\./);
assert.equal(/refreshes on open/i.test(html), false);
assert.equal(/refreshable/i.test(html), false);
assert.equal(/will refresh/i.test(html), false);
assert.equal(/automatically refresh/i.test(html), false);
assert.equal(html.includes('search volume has been verified'), false);
assert.equal((html.match(/<h1[\s>]/g) || []).length, 1);

const sitemap = readFileSync(new URL('../../public/sitemap.xml', import.meta.url), 'utf8');
const vercel = readFileSync(new URL('../../vercel.json', import.meta.url), 'utf8');
const tools = readFileSync(new URL('../../tools.html', import.meta.url), 'utf8');
const generator = readFileSync(new URL('../../scripts/generate-tool-pages.mjs', import.meta.url), 'utf8');
assert.match(sitemap, /https:\/\/www\.pivotchartmaker\.org\/export-excel-pivot-table</);
assert.match(vercel, /"source": "\/export-excel-pivot-table"/);
assert.match(tools, /href="\/export-excel-pivot-table"/);
assert.match(generator, /href="\/export-excel-pivot-table"/);

const officeDir = mkdtempSync(join(tmpdir(), 'pivot-office-'));
try {
  const xlsxPath = join(officeDir, 'book.xlsx');
  const outDir = join(officeDir, 'out');
  const profile = join(officeDir, 'profile');
  mkdirSync(outDir);
  writeFileSync(xlsxPath, built.bytes);
  execFileSync('soffice', [
    '--headless',
    '--norestore',
    '--nologo',
    '--nolockcheck',
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    '--convert-to',
    'ods',
    '--outdir',
    outDir,
    xlsxPath,
  ], { timeout: 60000, stdio: 'pipe' });
  const content = execFileSync('python3', ['-c', `
import sys, zipfile
print(zipfile.ZipFile(sys.argv[1]).read("content.xml").decode("utf-8"))
`, join(outDir, 'book.ods')], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 });
  assert.match(content, /table:data-pilot-table table:name="PivotTable1"/);
  assert.match(content, /table:cell-range-address="Source\.A1:Source\.C6"/);
  assert.match(content, /table:source-field-name="Region" table:orientation="row"/);
  assert.match(content, /table:source-field-name="Amount"[^>]*table:function="sum"/);
  assert.match(content, /table:grand-total="row"/);
} finally {
  rmSync(officeDir, { recursive: true, force: true });
}

console.log('excel pivot table export: all checks passed');
