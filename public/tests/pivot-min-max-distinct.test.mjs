import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { Buffer } from 'node:buffer';
import * as XLSX from '@e965/xlsx';
import {
  EXAMPLE_CSV,
  LIMITS,
  OVERALL_GROUP_LABEL,
  SCOPE_GROUP,
  SCOPE_OVERALL,
  aggregateMinMaxDistinct,
  assertByteSize,
  numericText,
  openXlsx,
  parseCsv,
  parseWorksheetXml,
  previewResult,
  serializeCsv,
  sharedStringsFromXml,
} from '../js/pivot-min-max-distinct.js';

function rowsOf(csv, group = 0, value = 1) {
  return aggregateMinMaxDistinct(parseCsv(csv), group, value);
}

const example = rowsOf(EXAMPLE_CSV);
assert.deepEqual(example.rows, [
  ['East', '1', '5', '2', SCOPE_GROUP],
  ['West', '5', '9', '3', SCOPE_GROUP],
  ['North', '2', '2', '2', SCOPE_GROUP],
  [OVERALL_GROUP_LABEL, '1', '9', '6', SCOPE_OVERALL],
]);
assert.equal(example.groupCount, 3);
assert.equal(example.dataRowCount, 8);
assert.equal(Number(example.rows[0][3]) + Number(example.rows[1][3]) + Number(example.rows[2][3]), 7);
assert.equal(example.rows[3][3], '6');

const blanks = rowsOf('Region,Amount\nEast,\nEast,5\nEast,\nWest, \nWest,5\n,3');
assert.deepEqual(blanks.rows, [
  ['East', '5', '5', '2', SCOPE_GROUP],
  ['West', '5', '5', '2', SCOPE_GROUP],
  ['', '3', '3', '1', SCOPE_GROUP],
  [OVERALL_GROUP_LABEL, '3', '5', '4', SCOPE_OVERALL],
]);

const duplicates = rowsOf('Team,Score\nA,1\nA,2\nA,1\na,2');
assert.deepEqual(duplicates.rows, [
  ['A', '1', '2', '2', SCOPE_GROUP],
  ['a', '2', '2', '1', SCOPE_GROUP],
  [OVERALL_GROUP_LABEL, '1', '2', '2', SCOPE_OVERALL],
]);

const mixed = rowsOf('Item,Value\nA,10\nA,n/a\nA,10\nA,foo\nA,-2\nB,$5\nB,5%\nB,"1,000"');
assert.deepEqual(mixed.rows, [
  ['A', '-2', '10', '4', SCOPE_GROUP],
  ['B', '', '', '3', SCOPE_GROUP],
  [OVERALL_GROUP_LABEL, '-2', '10', '7', SCOPE_OVERALL],
]);

const ties = rowsOf('G,V\nA,1.0\nA,1\nA,1e2\nA,50');
assert.deepEqual(ties.rows[0], ['A', '1.0', '1e2', '4', SCOPE_GROUP]);
assert.equal(numericText('1e309'), null);
assert.equal(numericText('  -0.5 ').text, '-0.5');
assert.equal(numericText('007').text, '007');

const spaced = rowsOf('G,V\nA, 5 \nA,5');
assert.deepEqual(spaced.rows[0], ['A', '5', '5', '2', SCOPE_GROUP]);

const sameColumn = rowsOf('Label\nEast\nEast\nWest', 0, 0);
assert.deepEqual(sameColumn.rows, [
  ['East', '', '', '1', SCOPE_GROUP],
  ['West', '', '', '1', SCOPE_GROUP],
  [OVERALL_GROUP_LABEL, '', '', '2', SCOPE_OVERALL],
]);

const collision = rowsOf('Group,Amount\n(all rows),1\n(all rows),3\nEast,2');
assert.equal(collision.rows.filter((row) => row[0] === '(all rows)' && row[4] === SCOPE_GROUP).length, 1);
assert.equal(collision.rows.at(-1)[4], SCOPE_OVERALL);
assert.equal(collision.rows.at(-1)[3], '3');

const quoted = rowsOf('Group,Amount\n"East, HQ",10\n"East, HQ",4\n"West ""A""",2');
assert.equal(quoted.rows[0][0], 'East, HQ');
assert.deepEqual(quoted.rows[0].slice(1, 4), ['4', '10', '2']);
assert.equal(quoted.rows[1][0], 'West "A"');

const semi = rowsOf('Region;Amount\nEast;1\nEast;4', 0, 1);
assert.equal(parseCsv('Region;Amount\nEast;1\nEast;4').delimiter, ';');
assert.deepEqual(semi.rows[0], ['East', '1', '4', '2', SCOPE_GROUP]);

const tabs = parseCsv('Region\tAmount\nEast\t8\nWest\t1');
assert.equal(tabs.delimiter, '\t');
assert.deepEqual(aggregateMinMaxDistinct(tabs, 0, 1).rows[0], ['East', '8', '8', '1', SCOPE_GROUP]);

const headers = parseCsv('Amount,Amount,\n1,2,3\n4,5,6');
assert.deepEqual(headers.headers, ['Amount', 'Amount (2)', 'Column 3']);

const short = rowsOf('Region,Amount\nEast');
assert.deepEqual(short.rows[0], ['East', '', '', '1', SCOPE_GROUP]);

assert.throws(() => parseCsv(''), /The file is empty/);
assert.throws(() => parseCsv(' \n\t'), /The file is empty/);
assert.throws(() => parseCsv('Region,Amount'), /no data rows/);
assert.throws(() => parseCsv('Region,Amount\n\n'), /no data rows/);
assert.throws(() => parseCsv('"open,quote'), /quote is not closed/);
assert.throws(() => aggregateMinMaxDistinct({ headers: ['A'], rows: [] }, 0, 0), /no data rows/);
assert.throws(() => aggregateMinMaxDistinct(parseCsv('A,B\n1,2'), -1, 0), /row field/);
assert.throws(() => aggregateMinMaxDistinct(parseCsv('A,B\n1,2'), 0, 5), /value field/);
assert.throws(() => aggregateMinMaxDistinct(parseCsv('A,B\n1,2'), 1.5, 0), /row field/);
assert.throws(() => assertByteSize(LIMITS.maxBytes + 1), /larger than 10 MB/);
assert.doesNotThrow(() => assertByteSize(LIMITS.maxBytes));

const wide = `${Array.from({ length: 1001 }, (_, i) => `c${i}`).join(',')}\n${Array.from({ length: 1001 }, () => '1').join(',')}`;
assert.throws(() => parseCsv(wide), /1,000 columns/);

const tall = ['G,V', ...Array.from({ length: 100001 }, () => 'A,1')].join('\n');
assert.throws(() => parseCsv(tall), /100,000 data rows/);

const manyGroups = aggregateMinMaxDistinct(parseCsv(['G,V', ...Array.from({ length: 13 }, (_, i) => `G${i},1`)].join('\n')), 0, 1);
const preview = previewResult(manyGroups);
assert.equal(preview.length, 13);
assert.equal(preview.at(-1)[4], SCOPE_OVERALL);
assert.equal(preview.filter((row) => row[4] === SCOPE_GROUP).length, 12);

const roundTrip = parseCsv(serializeCsv(example.headers, example.rows));
assert.deepEqual(roundTrip.headers, example.headers);
assert.deepEqual(roundTrip.rows, example.rows);
assert.equal(serializeCsv(['Group'], [['East, HQ']]).charCodeAt(0), 0xfeff);

const formulaXml = '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>East</t></is></c><c r="B2"><f>1+2</f><v>3</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>East</t></is></c><c r="B3" t="b"><v>1</v></c></row></sheetData></worksheet>';
const formulaRows = parseWorksheetXml(formulaXml, []);
assert.deepEqual(formulaRows, [['Region', 'Amount'], ['East', '3'], ['East', 'TRUE']]);

const rich = sharedStringsFromXml('<sst><si><r><t>Ea</t></r><r><t>st</t></r><rPh sb="0" eb="2"><t>SHOULD NOT APPEAR</t></rPh></si><si><t>a&amp;b</t></si></sst>');
assert.deepEqual(rich, ['East', 'a&b']);

function storeZip(files) {
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const [name, text] of files) {
    const nameBytes = Buffer.from(name);
    const data = Buffer.from(text);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 8);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    chunks.push(local, nameBytes, data);
    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0, 10);
    cd.writeUInt32LE(data.length, 20);
    cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nameBytes.length, 28);
    cd.writeUInt32LE(offset, 42);
    central.push(cd, nameBytes);
    offset += local.length + nameBytes.length + data.length;
  }
  const centralBuf = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(centralBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, centralBuf, eocd]);
}

const stored = storeZip([
  ['xl/workbook.xml', '<?xml version="1.0"?><workbook><sheets><sheet name="Sales &amp; &quot;Q1&quot;" sheetId="1" r:id="rId1"/></sheets></workbook>'],
  ['xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
  ['xl/sharedStrings.xml', '<sst><si><t>Region</t></si><si><t>Amount</t></si><si><r><t>Ea</t></r><r><t>st</t></r><rPh><t>NO</t></rPh></si><si><t></t></si></sst>'],
  ['xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>4</v></c></row><row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3" t="s"><v>3</v></c></row></sheetData></worksheet>'],
]);
const storedBook = await openXlsx(stored);
assert.deepEqual(storedBook.sheetNames, ['Sales & "Q1"']);
const storedSheet = await storedBook.sheet(0);
assert.equal(storedSheet.name, 'Sales & "Q1"');
assert.deepEqual(aggregateMinMaxDistinct(storedSheet, 0, 1).rows[0], ['East', '4', '4', '2', SCOPE_GROUP]);
await assert.rejects(storedBook.sheet(1), /worksheet that exists/);
await assert.rejects(openXlsx(Uint8Array.from([1, 2, 3, 4])), /could not be read/);

const inlineSheet = XLSX.utils.aoa_to_sheet([
  ['Region', 'Amount'],
  ['East', 5],
  ['East', 1],
  ['West', 'n/a'],
  ['North', null],
  ['North', 2],
]);
const inlineBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(inlineBook, inlineSheet, 'First');
XLSX.utils.book_append_sheet(inlineBook, XLSX.utils.aoa_to_sheet([['Region', 'Amount'], ['Only', 9]]), 'Second');
const inlineBytes = XLSX.write(inlineBook, { type: 'buffer', bookType: 'xlsx' });
const opened = await openXlsx(inlineBytes);
assert.deepEqual(opened.sheetNames, ['First', 'Second']);
const first = await opened.sheet(0);
assert.deepEqual(aggregateMinMaxDistinct(first, 0, 1).rows, [
  ['East', '1', '5', '2', SCOPE_GROUP],
  ['West', '', '', '1', SCOPE_GROUP],
  ['North', '2', '2', '2', SCOPE_GROUP],
  [OVERALL_GROUP_LABEL, '1', '5', '5', SCOPE_OVERALL],
]);
const second = await opened.sheet(1);
assert.deepEqual(aggregateMinMaxDistinct(second, 0, 1).rows[0], ['Only', '9', '9', '1', SCOPE_GROUP]);
const cached = await opened.sheet(0);
assert.equal(cached, first);

const sstSheet = XLSX.utils.aoa_to_sheet([
  ['Region', 'Amount'],
  ['East', 10],
  ['East', 'n/a'],
  ['East', 3],
]);
const sstBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(sstBook, sstSheet, 'S');
const sstBytes = XLSX.write(sstBook, { type: 'buffer', bookType: 'xlsx', bookSST: true });
const sstOpened = await openXlsx(sstBytes);
const sstTable = await sstOpened.sheet(0);
assert.deepEqual(aggregateMinMaxDistinct(sstTable, 0, 1).rows[0], ['East', '3', '10', '3', SCOPE_GROUP]);

const emptyBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(emptyBook, XLSX.utils.aoa_to_sheet([['Region', 'Amount']]), 'Empty');
await assert.rejects(openXlsx(XLSX.write(emptyBook, { type: 'buffer', bookType: 'xlsx' })).then((book) => book.sheet(0)), /no data rows/);

const html = readFileSync(new URL('../pivot-min-max-distinct.html', import.meta.url), 'utf8');
const title = html.match(/<title>([^<]*)<\/title>/)[1];
const description = html.match(/name="description" content="([^"]*)"/)[1];
assert.equal(title, 'Pivot Minimum, Maximum, Distinct Count');
assert.ok(title.length <= 60);
assert.ok((`${title} | Pivot Chart Maker`).length <= 60);
assert.ok(description.length >= 140 && description.length <= 160, `description length ${description.length}`);
assert.equal((html.match(/<h1[\s>]/g) || []).length, 1);
assert.match(html, /href="https:\/\/www\.pivotchartmaker\.org\/pivot-min-max-distinct"/);
assert.match(html, /distinct count 6/);
assert.doesNotMatch(html, /search volume|testimonial|stars|users say/i);

console.log('pivot min/max/distinct: all checks passed');
