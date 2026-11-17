import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as XLSX from '@e965/xlsx';
import {
  EXAMPLE_CSV,
  NA,
  RESULT_HEADERS,
  aggregatePercentChange,
  comparePeriods,
  formatDecimal,
  formatRatioPercent,
  measureNumber,
  openXlsx,
  parseCsv,
  parsePeriod,
  previewResult,
  serializeCsv,
  sortPeriods,
} from '../js/pivot-percent-change.js';

function rowsOf(csv, group = 0, period = 1, measure = 2) {
  return aggregatePercentChange(parseCsv(csv), group, period, measure);
}

function cells(result) {
  return result.rows.flat();
}

function assertNoInventedNumber(result) {
  for (const value of cells(result)) {
    assert.equal(value.includes('NaN'), false, value);
    assert.equal(value.includes('Infinity'), false, value);
  }
}

const example = rowsOf(EXAMPLE_CSV);
assert.deepEqual(example.headers, RESULT_HEADERS);
assert.deepEqual(example.periods, ['2024-01', '2024-02', '2024-04']);
assert.equal(example.grandTotal, '200');
assert.equal(example.periods.includes('2024-03'), false);
assert.deepEqual(example.rows, [
  ['East', '2024-01', '50', '25', '', NA, NA],
  ['East', '2024-02', '80', '40', '2024-01', '30', '60'],
  ['East', '2024-04', '20', '10', '2024-02', '-60', '-75'],
  ['West', '2024-01', '', NA, '', NA, NA],
  ['West', '2024-04', '50', '25', '2024-02', NA, NA],
  ['North', '2024-02', '0', '0', '2024-01', NA, NA],
]);
assert.equal(example.dataRowCount, 7);
assert.equal(example.groupCount, 3);
assertNoInventedNumber(example);

const gap = rowsOf('Region,Period,Amount\nEast,2024-01,10\nEast,2024-03,40');
assert.deepEqual(gap.periods, ['2024-01', '2024-03']);
assert.deepEqual(gap.rows, [
  ['East', '2024-01', '10', '20', '', NA, NA],
  ['East', '2024-03', '40', '80', '2024-01', '30', '300'],
]);
assertNoInventedNumber(gap);

const missingGroup = rowsOf('Region,Period,Amount\nWest,2024-01,10\nEast,2024-03,10');
assert.deepEqual(missingGroup.rows, [
  ['West', '2024-01', '10', '50', '', NA, NA],
  ['East', '2024-03', '10', '50', '2024-01', NA, NA],
]);
assert.equal(missingGroup.rows[1][5], NA);
assert.notEqual(missingGroup.rows[1][5], '0');
assert.notEqual(missingGroup.rows[1][5], '10');

const nulls = rowsOf('Region,Period,Amount\nEast,2024-01,\nEast,2024-01,\nEast,2024-02,5\nWest,2024-02,');
assert.deepEqual(nulls.rows, [
  ['East', '2024-01', '', NA, '', NA, NA],
  ['East', '2024-02', '5', '100', '2024-01', NA, NA],
  ['West', '2024-02', '', NA, '2024-01', NA, NA],
]);
assert.equal(nulls.grandTotal, '5');
assert.notEqual(nulls.rows[0][2], '0');
assert.notEqual(nulls.rows[1][5], '5');

const duplicates = rowsOf('Team,Period,Score\nA,2024-01,2\nA,2024-01,3\nA,2024-01,2\na,2024-01,4');
assert.deepEqual(duplicates.rows, [
  ['A', '2024-01', '7', '63.636364', '', NA, NA],
  ['a', '2024-01', '4', '36.363636', '', NA, NA],
]);
assert.equal(duplicates.groupCount, 2);

const spacedGroups = rowsOf('G,P,V\nEast,2024-01,1\nEast ,2024-01,3');
assert.equal(spacedGroups.groupCount, 2);
assert.deepEqual(spacedGroups.rows.map((row) => row[0]), ['East', 'East ']);

const trimmedPeriods = rowsOf('G,P,V\nEast, 2024-01 ,2\nEast,2024-01,3');
assert.equal(trimmedPeriods.rows.length, 1);
assert.deepEqual(trimmedPeriods.rows[0], ['East', '2024-01', '5', '100', '', NA, NA]);

const ignoredText = rowsOf('G,P,V\nEast,2024-01,$5\nEast,2024-01,5%\nEast,2024-01,"1,000"\nEast,2024-01,5\nEast,2024-02,n/a');
assert.deepEqual(ignoredText.rows, [
  ['East', '2024-01', '5', '100', '', NA, NA],
  ['East', '2024-02', '', NA, '2024-01', NA, NA],
]);

const zeroPrevious = rowsOf('G,P,V\nEast,2024-01,0\nEast,2024-02,5');
assert.deepEqual(zeroPrevious.rows, [
  ['East', '2024-01', '0', '0', '', NA, NA],
  ['East', '2024-02', '5', '100', '2024-01', '5', NA],
]);
assert.equal(zeroPrevious.grandTotalZero, false);

const allZeros = rowsOf('G,P,V\nEast,2024-01,0\nEast,2024-02,0');
assert.equal(allZeros.grandTotal, '0');
assert.equal(allZeros.grandTotalZero, true);
assert.deepEqual(allZeros.rows, [
  ['East', '2024-01', '0', NA, '', NA, NA],
  ['East', '2024-02', '0', NA, '2024-01', '0', NA],
]);
assertNoInventedNumber(allZeros);

const cancel = rowsOf('G,P,V\nEast,2024-01,5\nEast,2024-02,-5');
assert.equal(cancel.grandTotal, '0');
assert.deepEqual(cancel.rows, [
  ['East', '2024-01', '5', NA, '', NA, NA],
  ['East', '2024-02', '-5', NA, '2024-01', '-10', '-200'],
]);
assertNoInventedNumber(cancel);

const noNumbers = rowsOf('G,P,V\nEast,2024-01,\nEast,2024-02,n/a');
assert.equal(noNumbers.grandTotal, null);
assert.equal(noNumbers.grandTotalMissing, true);
assert.deepEqual(noNumbers.rows.map((row) => row[3]), [NA, NA]);
assert.deepEqual(noNumbers.rows.map((row) => row[5]), [NA, NA]);

const quarters = rowsOf('G,P,V\nEast,Q2,20\nEast,Q10,30\nEast,Q1,10');
assert.deepEqual(quarters.periods, ['Q1', 'Q10', 'Q2']);
assert.deepEqual(quarters.rows, [
  ['East', 'Q1', '10', '16.666667', '', NA, NA],
  ['East', 'Q10', '30', '50', 'Q1', '20', '200'],
  ['East', 'Q2', '20', '33.333333', 'Q10', '-10', '-33.333333'],
]);

const unpadded = rowsOf('G,P,V\nEast,2024-10,1\nEast,2024-2,1\nEast,2024-1,1');
assert.deepEqual(unpadded.periods, ['2024-1', '2024-2', '2024-10']);
assert.notDeepEqual(unpadded.periods, ['2024-1', '2024-10', '2024-2']);

const distinctLabels = rowsOf('G,P,V\nEast,2024-01,2\nEast,2024-1,3');
assert.equal(distinctLabels.rows.length, 2);
assert.deepEqual(distinctLabels.periods, ['2024-01', '2024-1']);
assert.equal(distinctLabels.rows[1][4], '2024-01');

const invalidDate = rowsOf('G,P,V\nEast,2024-02-31,4\nEast,2024-03-01,6');
assert.equal(parsePeriod('2024-02-31'), null);
assert.equal(parsePeriod('2023-02-29'), null);
assert.notEqual(parsePeriod('2024-02-29'), null);
assert.deepEqual(invalidDate.periods, ['2024-03-01', '2024-02-31']);
assert.equal(invalidDate.rows[1][4], '2024-03-01');

const zone = sortPeriods(['2024-01-02', '2024-01-02T00:00:00+08:00', '2024-01-01']);
assert.deepEqual(zone, ['2024-01-01', '2024-01-02T00:00:00+08:00', '2024-01-02']);
assert.equal(parsePeriod('2024-01-02T00:00:00Z'), Date.UTC(2024, 0, 2));
assert.equal(parsePeriod('2024-01-02T00:00:00.5Z'), Date.UTC(2024, 0, 2, 0, 0, 0, 500));
assert.equal(comparePeriods('Q1', '2024-01'), 1);

const thirds = rowsOf('G,P,V\nA,2024-01,1\nB,2024-01,2');
assert.deepEqual(thirds.rows.map((row) => row[3]), ['33.333333', '66.666667']);
assert.equal(formatRatioPercent(1, 6), '16.666667');
assert.equal(formatRatioPercent(2, 3), '66.666667');
assert.equal(formatRatioPercent(-10, 30), '-33.333333');
assert.equal(formatRatioPercent(1, 0), NA);
assert.equal(formatRatioPercent(1, null), NA);
assert.equal(formatDecimal(0.30000000000000004, 6), '0.3');
assert.equal(formatDecimal(1.0000005, 6), '1.000001');
assert.equal(formatDecimal(-75, 6), '-75');

const signed = rowsOf('G,P,V\nEast,2024-01,10\nEast,2024-02,-4');
assert.equal(signed.grandTotal, '6');
assert.deepEqual(signed.rows, [
  ['East', '2024-01', '10', '166.666667', '', NA, NA],
  ['East', '2024-02', '-4', '-66.666667', '2024-01', '-14', '-140'],
]);

const decimals = rowsOf('G,P,V\nEast,2024-01,0.1\nEast,2024-01,0.2');
assert.equal(decimals.rows[0][2], '0.3');
assert.equal(decimals.rows[0][3], '100');
assertNoInventedNumber(decimals);

const quoted = rowsOf('Group,Period,Amount\n"East, HQ",2024-01,10\n"East, HQ",2024-02,4');
assert.equal(quoted.rows[0][0], 'East, HQ');
assert.deepEqual(quoted.rows[1].slice(2), ['4', '28.571429', '2024-01', '-6', '-60']);

const semi = parseCsv('Region;Period;Amount\nEast;2024-01;2\nEast;2024-02;4');
assert.equal(semi.delimiter, ';');
assert.equal(aggregatePercentChange(semi, 0, 1, 2).rows[1][5], '2');

const negativeZero = rowsOf('G,P,V\nEast,2024-01,-0\nEast,2024-02,2');
assert.equal(negativeZero.rows[0][2], '0');
assert.equal(negativeZero.rows[1][5], '2');
assert.equal(negativeZero.rows[1][6], NA);

assert.equal(measureNumber(''), null);
assert.equal(measureNumber('n/a'), null);
assert.equal(measureNumber('007'), 7);
assert.throws(() => measureNumber('9007199254740993'), /totaled exactly/);
assert.throws(() => rowsOf(`G,P,V\nEast,2024-01,${Number.MAX_SAFE_INTEGER}\nEast,2024-01,1`), /too large/);

assert.throws(() => rowsOf('Only\nEast', 0, 0, 0), /different from the group field/);
assert.throws(() => aggregatePercentChange({ headers: ['A', 'B', 'C'], rows: [] }, 0, 1, 2), /no data rows/);
assert.throws(() => aggregatePercentChange(parseCsv('A,B,C\n1,2,3'), -1, 1, 2), /group field/);
assert.throws(() => aggregatePercentChange(parseCsv('A,B,C\n1,2,3'), 0, 1.5, 2), /period field/);
assert.throws(() => aggregatePercentChange(parseCsv('A,B,C\n1,2,3'), 0, 1, 9), /measure field/);
assert.throws(() => parseCsv(''), /The file is empty/);
assert.throws(() => parseCsv('"open'), /quote is not closed/);
assert.throws(() => parseCsv('A,B,C'), /no data rows/);

const many = rowsOf(['G,P,V', ...Array.from({ length: 13 }, (_, i) => `G${i},2024-01,1`)].join('\n'));
assert.equal(many.rows.length, 13);
assert.equal(previewResult(many).length, 12);
assert.equal(previewResult(example).length, example.rows.length);

const roundTrip = parseCsv(serializeCsv(example.headers, example.rows));
assert.deepEqual(roundTrip.headers, example.headers);
assert.deepEqual(roundTrip.rows, example.rows);
assert.equal(serializeCsv(example.headers, example.rows).charCodeAt(0), 0xfeff);
assert.deepEqual(rowsOf(`\uFEFF${EXAMPLE_CSV}`).rows, example.rows);

const sheet = XLSX.utils.aoa_to_sheet([
  ['Region', 'Period', 'Amount'],
  ['East', '2024-01', 10],
  ['East', '2024-03', 40],
  ['West', '2024-03', null],
]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, sheet, 'Sales');
const opened = await openXlsx(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
assert.deepEqual(opened.sheetNames, ['Sales']);
const xlsxRows = aggregatePercentChange(await opened.sheet(0), 0, 1, 2);
assert.deepEqual(xlsxRows.rows, [
  ['East', '2024-01', '10', '20', '', NA, NA],
  ['East', '2024-03', '40', '80', '2024-01', '30', '300'],
  ['West', '2024-03', '', NA, '2024-01', NA, NA],
]);
assertNoInventedNumber(xlsxRows);

const html = readFileSync(new URL('../pivot-percent-change.html', import.meta.url), 'utf8');
const title = html.match(/<title>([^<]*)<\/title>/)[1];
const description = html.match(/name="description" content="([^"]*)"/)[1];
const lede = html.match(/<p class="lede">([^<]*)<\/p>/)[1];
assert.equal(title, 'Pivot Percent of Total, Period Change');
assert.equal(title.length, 37);
assert.ok(title.length <= 60);
assert.equal((`${title} | Pivot Chart Maker`).length, 57);
assert.ok(description.length >= 140 && description.length <= 160, `description length ${description.length}`);
assert.notEqual(description, lede);
assert.equal((html.match(/<h1[\s>]/g) || []).length, 1);
assert.match(html, /href="https:\/\/www\.pivotchartmaker\.org\/pivot-percent-change"/);
assert.match(html, /2024-03 is not inserted/);
assert.match(html, /not a measurement of this site/);
assert.doesNotMatch(html, /search volume|testimonial|users say|star rating/i);

const sitemap = readFileSync(new URL('../sitemap.xml', import.meta.url), 'utf8');
assert.equal(sitemap.split('pivot-percent-change').length - 1, 1);
const vercel = readFileSync(new URL('../../vercel.json', import.meta.url), 'utf8');
assert.match(vercel, /"source": "\/pivot-percent-change"/);
const tools = readFileSync(new URL('../../tools.html', import.meta.url), 'utf8');
assert.match(tools, /href="\/pivot-percent-change"/);
const network = readFileSync(new URL('../../footer-network.links.json', import.meta.url), 'utf8');
assert.equal(network.includes('pivot-percent-change'), false);

console.log('pivot percent change: all checks passed');
