import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as XLSX from '@e965/xlsx';
import {
  EXAMPLE_COLUMN_NAME,
  EXAMPLE_CSV,
  EXAMPLE_FORMULA,
  calculateColumn,
  openXlsx,
  parseCsv,
  previewRows,
  serializeCsv,
} from '../js/calculated-columns.js';

function calc(csv, formula = EXAMPLE_FORMULA, name = EXAMPLE_COLUMN_NAME, delimiter) {
  return calculateColumn(parseCsv(csv, delimiter), formula, name);
}

const example = calc(EXAMPLE_CSV);
assert.deepEqual(example.headers, ['Item', 'Width', 'Height', 'Calculated']);
assert.equal(example.dataRowCount, 5);
assert.equal(example.calculatedCount, 3);
assert.equal(example.blankNotPlain, 2);
assert.equal(example.blankDivideByZero, 0);
assert.equal(example.blankNotFinite, 0);
assert.deepEqual(example.rows, [
  ['tile', '2', '3', '8'],
  ['panel', '10', '4', '18'],
  ['label', 'n/a', '1', ''],
  ['spacer', '', '5', ''],
  ['strip', '1.5', '2', '5.5'],
]);
assert.equal(JSON.stringify(example.rows).includes('NaN'), false);
assert.equal(JSON.stringify(example.rows).includes('Infinity'), false);
assert.equal(example.rows[2][3], '');
assert.notEqual(example.rows[2][3], '0');

const source = parseCsv(EXAMPLE_CSV);
const snapshot = structuredClone(source);
calculateColumn(source, EXAMPLE_FORMULA, EXAMPLE_COLUMN_NAME);
assert.deepEqual(source, snapshot);

assert.deepEqual(calc(EXAMPLE_CSV, '([Width] + [Height]) * 2').rows[0], ['tile', '2', '3', '10']);
assert.deepEqual(calc('Item,Width\na,8', '[Width] / 2 / 2', 'Calculated').rows, [['a', '8', '2']]);
assert.deepEqual(calc('Item,Width\na,4', '-[Width]', 'Calculated').rows, [['a', '4', '-4']]);
assert.deepEqual(calc('Item,Width\na,4', '+[Width]', 'Calculated').rows, [['a', '4', '4']]);
assert.deepEqual(calc('Item,Width\na,4', '--[Width]', 'Calculated').rows, [['a', '4', '4']]);
assert.deepEqual(calc('Item,Width,Height\na,2,3', '[Width] / [Height]').rows[0][3], '0.6666666666666666');
assert.deepEqual(calc('A,B\n0.1,0.2', '[A] + [B]', 'Calculated').rows, [['0.1', '0.2', '0.30000000000000004']]);

const divide = calc('Item,Width,Height\na,4,0\nb,n/a,0', '[Width] / [Height]');
assert.deepEqual(divide.rows, [
  ['a', '4', '0', ''],
  ['b', 'n/a', '0', ''],
]);
assert.equal(divide.blankDivideByZero, 1);
assert.equal(divide.blankNotPlain, 1);
assert.equal(JSON.stringify(divide.rows).includes('Infinity'), false);

assert.deepEqual(calc('Item,Width,Height\na,1,0', '[Width] / 0 + [Height]').rows, [['a', '1', '0', '']]);
assert.equal(calc('Item,Width,Height\na,1,0', '[Width] / 0 + [Height]').blankDivideByZero, 1);

const huge = calc('Item,Width\na,1e308', '[Width] * 10');
assert.deepEqual(huge.rows, [['a', '1e308', '']]);
assert.equal(huge.blankNotFinite, 1);
assert.equal(JSON.stringify(huge.rows).includes('Infinity'), false);

assert.deepEqual(calc('Item,Width\na,-0', '[Width] + 1').rows, [['a', '-0', '1']]);
assert.deepEqual(calc('Item,Left,Right\na,0,-1', '[Left] * [Right]', 'Calculated').rows[0][3], '0');
assert.deepEqual(calc('Item,Width\na, 10 ', '[Width] + 1').rows, [['a', ' 10 ', '11']]);
assert.deepEqual(calc('Item,Width\na,1e2', '[Width] + 0').rows, [['a', '1e2', '100']]);
assert.deepEqual(calc('Item,Width\na,10.0', '[Width] + 0').rows, [['a', '10.0', '10']]);
assert.deepEqual(calc('Item,Width\na,010', '[Width] + 1').rows, [['a', '010', '11']]);
assert.deepEqual(calc('Item,Width\na,1e309', '[Width] + 1').rows, [['a', '1e309', '']]);
assert.equal(calc('Item,Width\na,1e309', '[Width] + 1').blankNotPlain, 1);

for (const raw of ['', 'n/a', '$5', '5%', ' ']) {
  const row = calc(`Item,Width\na,${raw}`, '[Width] + 1').rows[0];
  assert.equal(row[2], '', raw);
  assert.notEqual(row[2], '1');
}
assert.deepEqual(calc('Item,Width\na,"1,000"', '[Width] + 1').rows, [['a', '1,000', '']]);

assert.deepEqual(calc('Item,Width\na,5', '[Width] - [Width]').rows, [['a', '5', '0']]);
assert.deepEqual(calc('2,3\n10,4', '[2] + 3', 'Calculated').rows, [['10', '4', '13']]);
assert.deepEqual(calc('Item,Box Width,Height\na,2,3', '[Box Width] * [Height]').rows, [['a', '2', '3', '6']]);
assert.deepEqual(calc('箱,边\n甲,2\n乙,3', '[边] * 2', 'Calculated').rows, [
  ['甲', '2', '4'],
  ['乙', '3', '6'],
]);

const renamed = calc('Width,Width\n2,3', '[Width (2)] + [Width]', 'Sum');
assert.deepEqual(renamed.headers, ['Width', 'Width (2)', 'Sum']);
assert.deepEqual(renamed.rows, [['2', '3', '5']]);

const blankHeader = calc(',Width\n7,1', '[Column 1] + [Width]', 'Calculated');
assert.deepEqual(blankHeader.headers, ['Column 1', 'Width', 'Calculated']);
assert.deepEqual(blankHeader.rows, [['7', '1', '8']]);

const bracket = calc('A]B,Width\n1,2', '[Width] + 1');
assert.deepEqual(bracket.rows, [['1', '2', '3']]);
assert.throws(() => calc('A]B,Width\n1,2', '[A]B] + 1'), /unexpected|no column named/);

assert.deepEqual(calc('Item,Width\na,4\nb,5', '  [Width]  +  1 ').rows.map((row) => row[2]), ['5', '6']);
assert.deepEqual(calc('Item,Width\na,4', '[Width] + 1', '  Calculated  ').headers[2], 'Calculated');

const short = calc('Item,Width,Height\na,4', '[Height] + 1');
assert.deepEqual(short.rows, [['a', '4', '', '']]);
assert.equal(short.blankNotPlain, 1);

assert.throws(() => calc(EXAMPLE_CSV, '1 + 2'), /at least one existing column/);
assert.throws(() => calc(EXAMPLE_CSV, '   '), /Enter a formula/);
assert.throws(() => calc(EXAMPLE_CSV, null), /Enter a formula/);
assert.throws(() => calc(EXAMPLE_CSV, '[width] + 1'), /no column named "width"/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width ] + 1'), /no column named "Width "/);
assert.throws(() => calc(EXAMPLE_CSV, '[Depth]'), /no column named "Depth"/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] + 1', 'Width'), /does not replace/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] + 1', '   '), /Enter a name/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] + 1', null), /Enter a name/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] + 1', `N${'a'.repeat(200)}`), /longer than 200/);
assert.throws(() => calc(EXAMPLE_CSV, `[Width]${'+'.repeat(2000)}1`), /longer than 2,000/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] + 1e309'), /too large/);
assert.throws(() => calc(EXAMPLE_CSV, 'SUM([Width])'), /unexpected "S"/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] > 1'), /unexpected ">"/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] ^ 2'), /unexpected "\^"/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width][Height]'), /does not multiply/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] +'), /ended before a value/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width] + * [Height]'), /missing a value/);
assert.throws(() => calc(EXAMPLE_CSV, '([Width]'), /missing a closing parenthesis/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width])'), /extra closing parenthesis/);
assert.throws(() => calc(EXAMPLE_CSV, '()'), /parentheses are empty/);
assert.throws(() => calc(EXAMPLE_CSV, '[]'), /empty/);
assert.throws(() => calc(EXAMPLE_CSV, '[Width'), /missing its closing/);
assert.throws(() => calc(EXAMPLE_CSV, `${'('.repeat(33)}[Width]${')'.repeat(33)}`), /too many nested parentheses/);
assert.deepEqual(calc(EXAMPLE_CSV, `${'('.repeat(32)}[Width]${')'.repeat(32)}`).rows[0][3], '2');

const many = ['Item,Width', ...Array.from({ length: 13 }, (_, index) => `r${index},${index}`)].join('\n');
const previewed = calc(many, '[Width] + 1');
assert.equal(previewed.rows.length, 13);
assert.equal(previewRows(previewed).length, 12);
assert.equal(previewRows(previewed)[0][2], '1');
assert.throws(() => previewRows(null), /no calculated column/);

const wideHeaders = Array.from({ length: 1000 }, (_, index) => `C${index}`);
assert.throws(() => calculateColumn({
  headers: wideHeaders,
  rows: [wideHeaders.map(() => '1')],
}, '[C0] + 1', 'Extra'), /cannot add another/);

const quoted = calc('Item,Width\n"a,b",2', '[Width] + 1', 'New,Name');
const csv = serializeCsv(quoted.headers, quoted.rows);
assert.equal(csv.charCodeAt(0), 0xfeff);
assert.equal(csv.includes('"a,b",2,3'), true);
assert.equal(csv.includes('"New,Name"'), true);
assert.equal(csv.includes('='), false);

const semi = calc('Item;Width\na;4', '[Width] + 1', 'Calculated', ';');
assert.deepEqual(semi.rows, [['a', '4', '5']]);

assert.throws(() => calculateColumn({ headers: [], rows: [['1']] }, '[A]', 'B'), /header row/);
assert.throws(() => calculateColumn({ headers: ['A'], rows: [] }, '[A]', 'B'), /no data rows/);
assert.throws(() => calculateColumn({ headers: ['A'], rows: [null] }, '[A] + 1', 'B'), /not a list of cells/);

const sheet = XLSX.utils.aoa_to_sheet([
  ['Item', 'Width', 'Height'],
  ['tile', 2, 3],
  ['label', 'n/a', 1],
]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, sheet, 'Boxes');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
  ['Item', 'Width', 'Height'],
  ['other', 4, 5],
]), 'Other');
const bytes = XLSX.write(book, { type: 'buffer', bookType: 'xlsx' });
const opened = await openXlsx(bytes);
assert.deepEqual(opened.sheetNames, ['Boxes', 'Other']);
const boxes = await opened.sheet(0);
const fromSheet = calculateColumn(boxes, EXAMPLE_FORMULA, EXAMPLE_COLUMN_NAME);
assert.equal(boxes.name, 'Boxes');
assert.deepEqual(fromSheet.rows, [
  ['tile', '2', '3', '8'],
  ['label', 'n/a', '1', ''],
]);
const other = await opened.sheet(1);
assert.deepEqual(calculateColumn(other, '[Width] * [Height]', 'Area').rows, [['other', '4', '5', '20']]);
await assert.rejects(openXlsx(new TextEncoder().encode('this is not a workbook')), /could not be read/);

const engine = readFileSync(new URL('../js/calculated-columns.js', import.meta.url), 'utf8');
assert.equal(engine.includes('eval('), false);
assert.equal(engine.includes('new Function'), false);

const html = readFileSync(new URL('../calculated-columns.html', import.meta.url), 'utf8');
const title = html.match(/<title>([^<]*)<\/title>/)[1];
const description = html.match(/<meta name="description" content="([^"]*)"/)[1];
const lede = html.match(/<p class="lede">([^<]*)<\/p>/)[1];
assert.equal(title, 'Calculated CSV Column');
assert.equal([...html.matchAll(/<h1>/g)].length, 1);
assert.equal(html.includes('<h1>Calculated CSV Column</h1>'), true);
assert.equal(title.length + ' | Pivot Chart Maker'.length <= 60, true);
assert.equal(description.length >= 140 && description.length <= 160, true, String(description.length));
assert.notEqual(description, lede);
assert.match(html, /https:\/\/www\.pivotchartmaker\.org\/calculated-columns/);
assert.match(html, /\[Width\] \+ \[Height\] \* 2/);
assert.match(html, /0\.30000000000000004/);
assert.match(html, /0\.6666666666666666/);
assert.match(html, /not an Excel formula/);
assert.match(html, /Search volume for this page has not been verified/);
assert.equal(html.includes('search volume has been verified'), false);
assert.equal(/\b(ROI|credit score|profit|revenue)\b/i.test(html), false);

const faqs = [...html.matchAll(/<summary>([^<]*)<\/summary>\s*<p>([^<]*)<\/p>/g)];
const schema = JSON.parse(html.match(/<script type="application\/ld\+json">([^<]*)<\/script>/)[1]);
const entities = schema['@graph'].find((item) => item['@type'] === 'FAQPage').mainEntity;
assert.equal(entities.length, faqs.length);
assert.equal(entities.length, 6);
for (let index = 0; index < faqs.length; index += 1) {
  assert.equal(entities[index].name, faqs[index][1]);
  assert.equal(entities[index].acceptedAnswer.text, faqs[index][2]);
}

const root = new URL('../../', import.meta.url);
const tools = readFileSync(new URL('tools.html', root), 'utf8');
const generator = readFileSync(new URL('scripts/generate-tool-pages.mjs', root), 'utf8');
const sitemap = readFileSync(new URL('public/sitemap.xml', root), 'utf8');
const vercel = readFileSync(new URL('vercel.json', root), 'utf8');
for (const text of [tools, generator]) {
  assert.match(text, /href="\/calculated-columns"/);
}
assert.match(sitemap, /https:\/\/www\.pivotchartmaker\.org\/calculated-columns/);
assert.match(vercel, /"source": "\/calculated-columns"/);
assert.match(vercel, /"destination": "\/calculated-columns\.html"/);

console.log('calculated columns: all checks passed');
