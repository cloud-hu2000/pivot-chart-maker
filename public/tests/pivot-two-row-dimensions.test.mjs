import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as XLSX from '@e965/xlsx';
import {
  BASE_HEADERS,
  COLUMN_HEADERS,
  EXAMPLE_CSV,
  OVERALL_LABEL,
  SCOPE_GROUP,
  SCOPE_OVERALL,
  aggregateTwoRowPivot,
  openXlsx,
  parseCsv,
  previewResult,
  serializeCsv,
} from '../js/pivot-two-row-dimensions.js';

function rowsOf(csv, row1 = 0, row2 = 1, value = 2, column = null) {
  return aggregateTwoRowPivot(parseCsv(csv), row1, row2, value, column);
}

function assertNoInventedNumber(result) {
  for (const value of result.rows.flat()) {
    assert.equal(String(value).includes('NaN'), false, value);
    assert.equal(String(value).includes('Infinity'), false, value);
  }
}

const example = rowsOf(EXAMPLE_CSV, 0, 1, 3, null);
assert.deepEqual(example.headers, BASE_HEADERS);
assert.equal(example.columnField, false);
assert.equal(example.pairCount, 6);
assert.equal(example.dataRowCount, 11);
assert.deepEqual(example.rows, [
  ['East', 'Desk', '22', '7.333333', '3', SCOPE_GROUP],
  ['East', 'Chair', '8', '8', '1', SCOPE_GROUP],
  ['West', 'Desk', '4', '4', '1', SCOPE_GROUP],
  ['West', 'Chair', '12', '6', '2', SCOPE_GROUP],
  ['North', 'Desk', '3', '3', '1', SCOPE_GROUP],
  ['east', 'Desk', '1', '1', '1', SCOPE_GROUP],
  [OVERALL_LABEL, '', '50', '5.555556', '9', SCOPE_OVERALL],
]);
assert.equal(example.sum, '50');
assert.equal(example.average, '5.555556');
assertNoInventedNumber(example);
assert.equal(JSON.stringify(example.rows).includes('West","Chair","Store'), false);

const withColumn = rowsOf(EXAMPLE_CSV, 0, 1, 3, 2);
assert.deepEqual(withColumn.headers, COLUMN_HEADERS);
assert.equal(withColumn.columnField, true);
assert.equal(withColumn.pairCount, 6);
assert.deepEqual(withColumn.rows, [
  ['East', 'Desk', 'Online', '15', '7.5', '2', SCOPE_GROUP],
  ['East', 'Desk', 'Store', '7', '7', '1', SCOPE_GROUP],
  ['East', 'Chair', 'Online', '8', '8', '1', SCOPE_GROUP],
  ['East', 'Chair', 'Store', '', '', '0', SCOPE_GROUP],
  ['West', 'Desk', 'Online', '4', '4', '1', SCOPE_GROUP],
  ['West', 'Desk', 'Store', '', '', '0', SCOPE_GROUP],
  ['West', 'Chair', 'Online', '12', '6', '2', SCOPE_GROUP],
  ['North', 'Desk', 'Store', '3', '3', '1', SCOPE_GROUP],
  ['east', 'Desk', 'Online', '1', '1', '1', SCOPE_GROUP],
  [OVERALL_LABEL, '', '', '50', '5.555556', '9', SCOPE_OVERALL],
]);
assert.equal(withColumn.rows.some((row) => row[0] === 'West' && row[1] === 'Chair' && row[2] === 'Store'), false);
assertNoInventedNumber(withColumn);
assert.equal(rowsOf(EXAMPLE_CSV, 0, 1, 3, '').columnField, false);
assert.equal(rowsOf(EXAMPLE_CSV, 0, 1, 3, undefined).columnField, false);

const nulls = rowsOf('Region,Product,Amount\nEast,Desk,\nEast,Desk,\nEast,Chair,5\nWest,Desk,');
assert.deepEqual(nulls.rows, [
  ['East', 'Desk', '', '', '0', SCOPE_GROUP],
  ['East', 'Chair', '5', '5', '1', SCOPE_GROUP],
  ['West', 'Desk', '', '', '0', SCOPE_GROUP],
  [OVERALL_LABEL, '', '5', '5', '1', SCOPE_OVERALL],
]);
assert.notEqual(nulls.rows[0][2], '0');
assert.equal(nulls.rows[0][4], '0');

const duplicates = rowsOf('Region,Product,Amount\nEast,Desk,2\nEast,Desk,3\nEast,Desk,2\neast,Desk,4\nEast,Desk ,9');
assert.equal(duplicates.pairCount, 3);
assert.deepEqual(duplicates.rows, [
  ['East', 'Desk', '7', '2.333333', '3', SCOPE_GROUP],
  ['east', 'Desk', '4', '4', '1', SCOPE_GROUP],
  ['East', 'Desk ', '9', '9', '1', SCOPE_GROUP],
  [OVERALL_LABEL, '', '20', '4', '5', SCOPE_OVERALL],
]);

const spacedOuter = rowsOf('Region,Product,Amount\nEast,Desk,1\nEast ,Desk,3');
assert.equal(spacedOuter.pairCount, 2);
assert.deepEqual(spacedOuter.rows.map((row) => row[0]), ['East', 'East ', OVERALL_LABEL]);

const shortRow = rowsOf('Region,Product,Amount\nEast,Desk,2\nEast');
assert.deepEqual(shortRow.rows, [
  ['East', 'Desk', '2', '2', '1', SCOPE_GROUP],
  ['East', '', '', '', '0', SCOPE_GROUP],
  [OVERALL_LABEL, '', '2', '2', '1', SCOPE_OVERALL],
]);
assert.equal(shortRow.rows.some((row) => row[1] === '(blank)' || row[1] === '(all)' || row[1] === '(missing)'), false);

const emptySecondCells = rowsOf('Region,Product,Amount\nEast,,3\n,Desk,4');
assert.deepEqual(emptySecondCells.rows, [
  ['East', '', '3', '3', '1', SCOPE_GROUP],
  ['', 'Desk', '4', '4', '1', SCOPE_GROUP],
  [OVERALL_LABEL, '', '7', '3.5', '2', SCOPE_OVERALL],
]);

assert.throws(
  () => aggregateTwoRowPivot(parseCsv('Region,Amount\nEast,5\nEast,1'), 0, '', 1),
  /does not invent a second row dimension/,
);
assert.throws(
  () => aggregateTwoRowPivot(parseCsv('Region,Amount\nEast,5'), 0, null, 1),
  /does not invent a second row dimension/,
);
assert.throws(
  () => aggregateTwoRowPivot(parseCsv('Region,Product,Amount\nEast,Desk,5'), 0, 0, 2),
  /different from the first/,
);
assert.throws(
  () => aggregateTwoRowPivot(parseCsv('Region,Product,Amount\nEast,Desk,5'), 0, 1, 2, 0),
  /different from both row fields/,
);

const ignored = rowsOf('Region,Product,Amount\nEast,Desk,$5\nEast,Desk,5%\nEast,Desk,"1,000"\nEast,Desk,5\nEast,Chair,n/a\nEast,Lamp,1e2');
assert.deepEqual(ignored.rows, [
  ['East', 'Desk', '5', '5', '1', SCOPE_GROUP],
  ['East', 'Chair', '', '', '0', SCOPE_GROUP],
  ['East', 'Lamp', '100', '100', '1', SCOPE_GROUP],
  [OVERALL_LABEL, '', '105', '52.5', '2', SCOPE_OVERALL],
]);

const zeros = rowsOf('Region,Product,Amount\nEast,Desk,0\nEast,Desk,-0\nEast,Chair,5\nEast,Chair,-5');
assert.deepEqual(zeros.rows, [
  ['East', 'Desk', '0', '0', '2', SCOPE_GROUP],
  ['East', 'Chair', '0', '0', '2', SCOPE_GROUP],
  [OVERALL_LABEL, '', '0', '0', '4', SCOPE_OVERALL],
]);

const stringIndexes = aggregateTwoRowPivot(parseCsv(EXAMPLE_CSV), '0', '1', '3', '');
assert.equal(stringIndexes.columnField, false);
assert.equal(stringIndexes.sum, '50');

const negative = rowsOf('Region,Product,Amount\nEast,Desk,-5\nEast,Desk,2');
assert.deepEqual(negative.rows[0], ['East', 'Desk', '-3', '-1.5', '2', SCOPE_GROUP]);

const notAverageOfAverages = rowsOf('Region,Product,Amount\nEast,Desk,1\nEast,Desk,1\nWest,Chair,10');
assert.equal(notAverageOfAverages.average, '4');
assert.notEqual(notAverageOfAverages.average, '5.5');
assert.deepEqual(notAverageOfAverages.rows.at(-1), [OVERALL_LABEL, '', '12', '4', '3', SCOPE_OVERALL]);

const huge = '9'.repeat(16);
assert.throws(() => rowsOf(`Region,Product,Amount\nEast,Desk,${huge}`), /totaled exactly/);
assert.throws(
  () => rowsOf('Region,Product,Amount\nEast,Desk,9007199254740991\nEast,Desk,1'),
  /too large to calculate exactly/,
);
assert.equal(rowsOf('Region,Product,Amount\nEast,Desk,1e309\nEast,Chair,4').rows[0][2], '');
assert.equal(rowsOf('Region,Product,Amount\nEast,Desk,1e309\nEast,Chair,4').sum, '4');

const decimals = rowsOf('Region,Product,Amount\nEast,Desk,1.5\nEast,Desk,2.25');
assert.deepEqual(decimals.rows[0], ['East', 'Desk', '3.75', '1.875', '2', SCOPE_GROUP]);

const missingIntersection = rowsOf('Region,Product,Channel,Amount\nA,X,Online,1\nA,Y,Store,2', 0, 1, 3, 2);
assert.equal(missingIntersection.detailCount, 2);
assert.deepEqual(missingIntersection.rows.map((row) => row.slice(0, 3)), [
  ['A', 'X', 'Online'],
  ['A', 'Y', 'Store'],
  [OVERALL_LABEL, '', ''],
]);

const many = ['Region,Product,Amount'];
for (let i = 0; i < 13; i += 1) many.push(`R${i},Item,${i + 1}`);
const wide = rowsOf(many.join('\n'));
assert.equal(wide.detailCount, 13);
const preview = previewResult(wide);
assert.equal(preview.length, 13);
assert.equal(preview[0][0], 'R0');
assert.equal(preview[11][0], 'R11');
assert.equal(preview.at(-1)[0], OVERALL_LABEL);
assert.equal(preview.some((row) => row[0] === 'R12'), false);

const namedOverall = rowsOf('Region,Product,Amount\n(all rows),Desk,6\nEast,Chair,4');
assert.equal(namedOverall.rows[0][0], OVERALL_LABEL);
assert.equal(namedOverall.rows[0][5], SCOPE_GROUP);
assert.equal(namedOverall.rows.at(-1)[5], SCOPE_OVERALL);
assert.equal(namedOverall.rows.at(-1)[2], '10');

const csv = serializeCsv(example.headers, example.rows);
assert.match(csv, /East,Desk,22,7\.333333,3,group/);
assert.match(csv, /\(all rows\),,50,5\.555556,9,all rows/);

const sheet = XLSX.utils.aoa_to_sheet([
  ['Region', 'Product', 'Amount'],
  ['East', 'Desk', 10],
  ['East', 'Chair', 5],
  ['West', 'Desk', 'n/a'],
]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, sheet, 'Sales');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Region', 'Product', 'Amount'], ['Only', 'Lamp', 9]]), 'Second');
const opened = await openXlsx(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
assert.deepEqual(opened.sheetNames, ['Sales', 'Second']);
const first = await opened.sheet(0);
assert.deepEqual(aggregateTwoRowPivot(first, 0, 1, 2).rows, [
  ['East', 'Desk', '10', '10', '1', SCOPE_GROUP],
  ['East', 'Chair', '5', '5', '1', SCOPE_GROUP],
  ['West', 'Desk', '', '', '0', SCOPE_GROUP],
  [OVERALL_LABEL, '', '15', '7.5', '2', SCOPE_OVERALL],
]);
const second = await opened.sheet(1);
assert.deepEqual(aggregateTwoRowPivot(second, 0, 1, 2).rows[0], ['Only', 'Lamp', '9', '9', '1', SCOPE_GROUP]);

const html = readFileSync(new URL('../pivot-two-row-dimensions.html', import.meta.url), 'utf8');
const title = html.match(/<title>([^<]*)<\/title>/)[1];
const description = html.match(/name="description" content="([^"]*)"/)[1];
const h1 = html.match(/<h1>([^<]*)<\/h1>/)[1];
assert.equal(title, 'Two Row Pivot Sum, Average, Count');
assert.ok(title.length <= 60);
assert.ok((`${title} | Pivot Chart Maker`).length <= 60, title.length);
assert.ok(description.length >= 140 && description.length <= 160, `description length ${description.length}`);
assert.equal((html.match(/<h1[\s>]/g) || []).length, 1);
assert.match(h1, /Pivot Two Row Dimensions/);
assert.match(description, /two row dimensions/);
assert.notEqual(description, html.match(/<p class="lede">([^<]*)<\/p>/)[1]);
assert.match(html, /href="https:\/\/www\.pivotchartmaker\.org\/pivot-two-row-dimensions"/);
assert.match(html, /7\.333333/);
assert.match(html, /5\.555556/);
assert.match(html, /does not invent a second row dimension/);
assert.match(html, /no West \/ Chair \/ Store row/);
assert.doesNotMatch(html, /search volume|testimonial|stars|users say|ranking/i);

console.log('pivot two row dimensions: all checks passed');
