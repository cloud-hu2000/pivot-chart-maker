import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as XLSX from '@e965/xlsx';
import {
  EXAMPLE_CONDITIONS,
  EXAMPLE_CSV,
  filterRows,
  openXlsx,
  parseCsv,
  previewMatches,
  serializeCsv,
} from '../js/csv-filter.js';

function rowsOf(csv, conditions, mode = 'and') {
  return filterRows(parseCsv(csv), conditions, mode);
}

const example = rowsOf(EXAMPLE_CSV, EXAMPLE_CONDITIONS, 'and');
assert.deepEqual(example.headers, ['Region', 'Product', 'Amount']);
assert.equal(example.matchCount, 1);
assert.equal(example.dataRowCount, 6);
assert.deepEqual(example.rows, [['East', 'Desk', '10']]);
assert.equal(JSON.stringify(example.rows).includes('Chair'), false);
assert.equal(JSON.stringify(example.rows).includes('West'), false);
assert.equal(JSON.stringify(example.rows).includes('NaN'), false);

const either = rowsOf(EXAMPLE_CSV, [
  { column: 0, operator: 'equals', value: 'West' },
  { column: 2, operator: 'equals', value: '10' },
], 'or');
assert.deepEqual(either.rows, [
  ['East', 'Desk', '10'],
  ['West', 'Desk', 'n/a'],
  ['North', 'Desk', '10'],
]);
assert.equal(either.rows.some((row) => row[0] === 'west'), false);
assert.equal(either.rows.some((row) => row[1] === 'Chair'), false);

const text = rowsOf([
  'Name,Note',
  'East,Desk',
  'east,desk',
  'East ,Desk',
  'West,Desk lamp',
  'North,Lamp',
].join('\n'), [
  { column: 0, operator: 'equals', value: 'East' },
]);
assert.deepEqual(text.rows, [['East', 'Desk']]);

assert.deepEqual(rowsOf('Name,Note\nEast,Desk\neast,desk', [
  { column: 0, operator: 'not_equals', value: 'East' },
]).rows, [['east', 'desk']]);

assert.deepEqual(rowsOf('Name,Note\nEast,Desk\nWest,Lamp', [
  { column: 1, operator: 'contains', value: 'e' },
]).rows, [['East', 'Desk']]);

assert.deepEqual(rowsOf('Name,Note\nEast,Desk\nWest,Lamp', [
  { column: 1, operator: 'not_contains', value: 'e' },
]).rows, [['West', 'Lamp']]);

assert.deepEqual(rowsOf('Name,Note\nEast,Desk\nWest,Lamp', [
  { column: 1, operator: 'starts_with', value: 'De' },
]).rows, [['East', 'Desk']]);

assert.deepEqual(rowsOf('Name,Note\nEast,Desk\nWest,Lamp', [
  { column: 1, operator: 'ends_with', value: 'mp' },
]).rows, [['West', 'Lamp']]);

const blanks = rowsOf('Region,Amount\nEast,\nEast, \nWest,0\nNorth,n/a', [
  { column: 1, operator: 'empty', value: '' },
]);
assert.deepEqual(blanks.rows, [['East', '']]);

assert.deepEqual(rowsOf('Region,Amount\nEast,\nWest,0', [
  { column: 1, operator: 'not_empty' },
]).rows, [['West', '0']]);

const numbers = rowsOf([
  'Label,Amount',
  'a,10',
  'b,6',
  'c,6.5',
  'd,-2',
  'e,n/a',
  'f,',
  'g,$5',
  'h,"1,000"',
  'i, 10 ',
  'j,1e2',
  'k,1e309',
].join('\n'), [
  { column: 1, operator: 'gt', value: '6' },
]);
assert.deepEqual(numbers.rows.map((row) => row[0]), ['a', 'c', 'i', 'j']);

assert.deepEqual(rowsOf('Label,Amount\na,10\nb,6\nc,4', [
  { column: 1, operator: 'gte', value: '6' },
]).rows.map((row) => row[0]), ['a', 'b']);

assert.deepEqual(rowsOf('Label,Amount\na,10\nb,6\nc,4', [
  { column: 1, operator: 'lt', value: '6' },
]).rows.map((row) => row[0]), ['c']);

assert.deepEqual(rowsOf('Label,Amount\na,10\nb,6\nc,4', [
  { column: 1, operator: 'lte', value: '6' },
]).rows.map((row) => row[0]), ['b', 'c']);

assert.deepEqual(rowsOf('Label,Amount\na,-0\nb,0\nc,1', [
  { column: 1, operator: 'lte', value: '0' },
]).rows.map((row) => row[0]), ['a', 'b']);

assert.throws(() => rowsOf('Label,Amount\na,5', [
  { column: 1, operator: 'gt', value: 'n/a' },
]), /not a plain number/);

assert.throws(() => rowsOf(EXAMPLE_CSV, [], 'and'), /does not keep every row/);
assert.throws(() => rowsOf(EXAMPLE_CSV, null, 'or'), /does not keep every row/);
assert.throws(() => rowsOf(EXAMPLE_CSV, EXAMPLE_CONDITIONS, 'xor'), /does not mix both/);
assert.throws(() => rowsOf(EXAMPLE_CSV, [{ column: 9, operator: 'equals', value: 'East' }]), /Choose a column/);
assert.throws(() => rowsOf(EXAMPLE_CSV, [{ column: 0, operator: 'equals', value: '' }]), /is empty/);
assert.throws(() => rowsOf(EXAMPLE_CSV, [{ column: 0, operator: 'regex', value: 'East' }]), /Choose a filter operator/);

const none = rowsOf(EXAMPLE_CSV, [{ column: 0, operator: 'equals', value: 'Nowhere' }]);
assert.equal(none.matchCount, 0);
assert.deepEqual(none.rows, []);
assert.deepEqual(none.headers, ['Region', 'Product', 'Amount']);
const noneCsv = serializeCsv(none.headers, none.rows);
assert.equal(noneCsv, '\uFEFFRegion,Product,Amount\n');
assert.equal(noneCsv.includes('Nowhere'), false);

const quoted = rowsOf('Name,Note\n"East, desk","keep, me"\nWest,drop', [
  { column: 0, operator: 'contains', value: 'desk' },
]);
assert.deepEqual(quoted.rows, [['East, desk', 'keep, me']]);

const short = rowsOf('Region,Product,Amount\nEast\nEast,Desk,10', [
  { column: 1, operator: 'empty' },
]);
assert.deepEqual(short.rows, [['East', '', '']]);

const both = rowsOf('Region,Amount\nEast,10\nEast,4\nWest,10', [
  { column: 0, operator: 'equals', value: 'East' },
  { column: 1, operator: 'gte', value: '10' },
], 'and');
assert.deepEqual(both.rows, [['East', '10']]);

const many = Array.from({ length: 15 }, (_, index) => `r${index},1`).join('\n');
const preview = rowsOf(`Name,Amount\n${many}`, [{ column: 1, operator: 'equals', value: '1' }]);
assert.equal(preview.matchCount, 15);
assert.equal(previewMatches(preview).length, 12);
assert.equal(previewMatches(preview)[0][0], 'r0');
assert.equal(preview.rows[14][0], 'r14');

const downloaded = serializeCsv(example.headers, example.rows);
assert.match(downloaded, /East,Desk,10/);
assert.equal(downloaded.includes('Chair'), false);
assert.equal(downloaded.includes('North'), false);

const sheet = XLSX.utils.aoa_to_sheet([
  ['Region', 'Product', 'Amount'],
  ['East', 'Desk', 10],
  ['East', 'Chair', 5],
  ['West', 'Desk', 'n/a'],
]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, sheet, 'Sales');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
  ['Region', 'Product', 'Amount'],
  ['North', 'Desk', 3],
]), 'Other');
const bytes = XLSX.write(book, { type: 'buffer', bookType: 'xlsx' });
const opened = await openXlsx(bytes);
assert.deepEqual(opened.sheetNames, ['Sales', 'Other']);
const sales = await opened.sheet(0);
const fromSheet = filterRows(sales, EXAMPLE_CONDITIONS, 'and');
assert.equal(sales.name, 'Sales');
assert.deepEqual(fromSheet.rows, [['East', 'Desk', '10']]);
const other = await opened.sheet(1);
assert.deepEqual(filterRows(other, [{ column: 0, operator: 'equals', value: 'North' }], 'and').rows, [['North', 'Desk', '3']]);

const html = readFileSync(new URL('../csv-filter.html', import.meta.url), 'utf8');
const title = html.match(/<title>([^<]*)<\/title>/)[1];
const description = html.match(/<meta name="description" content="([^"]*)"/)[1];
const lede = html.match(/<p class="lede">([^<]*)<\/p>/)[1];
assert.equal(title, 'Filter CSV by Column Conditions');
assert.equal([...html.matchAll(/<h1>/g)].length, 1);
assert.equal(html.includes('<h1>Filter CSV by Column Conditions</h1>'), true);
assert.equal(title.length + ' | Pivot Chart Maker'.length <= 60, true);
assert.equal(description.length >= 140 && description.length <= 160, true, description.length);
assert.notEqual(description, lede);
assert.match(html, /https:\/\/www\.pivotchartmaker\.org\/csv-filter/);
assert.match(html, /do not apply regular expressions/);
assert.match(html, /matching rows only/);
assert.equal(html.includes('search volume has been verified'), false);

console.log('csv filter: all checks passed');
