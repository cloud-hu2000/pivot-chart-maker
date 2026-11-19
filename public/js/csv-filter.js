/**
 * Filter CSV or XLSX rows by column conditions.
 *
 * One mode applies to the whole list. AND keeps a row only when every
 * condition matches. OR keeps a row when any one condition matches. There
 * are no parentheses and no mixed groups. With no conditions, this function
 * throws. It does not return every row.
 *
 * Text operators compare the cell text exactly. They do not trim, ignore
 * letter case, or apply regular expressions. "East" and "east" are different,
 * and "East " is not "East".
 *
 * Number operators use numericText: optional sign, decimal, or exponent,
 * after trimming the cell and the comparison value. Blank cells, words,
 * currency symbols, percent signs, and thousands separators do not match.
 * They are not treated as zero. A comparison value that is not a plain
 * number throws instead of matching nothing silently.
 *
 * "is empty" matches only the empty string. A cell that contains spaces is
 * not empty. Missing cells on a short row are the empty string.
 *
 * Matching rows stay in file order. The result does not add rows that were
 * not in the source. Files are read locally by the caller.
 */

import { numericText } from './pivot-min-max-distinct.js';

export {
  LIMITS,
  assertByteSize,
  openXlsx,
  parseCsv,
  serializeCsv,
} from './pivot-min-max-distinct.js';

export const PREVIEW_ROW_LIMIT = 12;

export const OPERATORS = [
  { id: 'equals', label: 'equals', kind: 'text' },
  { id: 'not_equals', label: 'does not equal', kind: 'text' },
  { id: 'contains', label: 'contains', kind: 'text' },
  { id: 'not_contains', label: 'does not contain', kind: 'text' },
  { id: 'starts_with', label: 'starts with', kind: 'text' },
  { id: 'ends_with', label: 'ends with', kind: 'text' },
  { id: 'gt', label: 'greater than', kind: 'number' },
  { id: 'gte', label: 'greater than or equal', kind: 'number' },
  { id: 'lt', label: 'less than', kind: 'number' },
  { id: 'lte', label: 'less than or equal', kind: 'number' },
  { id: 'empty', label: 'is empty', kind: 'empty' },
  { id: 'not_empty', label: 'is not empty', kind: 'empty' },
];

const OPERATOR_IDS = new Set(OPERATORS.map((item) => item.id));

export const EXAMPLE_CSV = [
  'Region,Product,Amount',
  'East,Desk,10',
  'East,Chair,5',
  'West,Desk,n/a',
  'west,Desk,4',
  'East,Lamp,',
  'North,Desk,10',
].join('\n');

export const EXAMPLE_CONDITIONS = [
  { column: 0, operator: 'equals', value: 'East' },
  { column: 2, operator: 'gt', value: '6' },
];

function cellString(value) {
  return value == null ? '' : String(value);
}

function requireTable(table) {
  if (!table || !Array.isArray(table.headers) || table.headers.length === 0) {
    throw new Error('A header row is required.');
  }
  if (!Array.isArray(table.rows)) throw new Error('Data rows are missing.');
  if (table.rows.length === 0) throw new Error('The file has a header but no data rows.');
  return table.headers.length;
}

function columnIndex(condition, width) {
  const index = Number(condition && condition.column);
  if (!Number.isInteger(index) || index < 0 || index >= width) {
    throw new Error('Choose a column for each condition.');
  }
  return index;
}

function comparisonValue(condition) {
  if (condition.value == null) return '';
  return String(condition.value);
}

function plainNumber(raw, label) {
  const parsed = numericText(raw);
  if (!parsed) {
    throw new Error(`${label} is not a plain number. Use a value such as 10, -2, or 1e2.`);
  }
  return Object.is(parsed.value, -0) ? 0 : parsed.value;
}

function textMatches(cell, operator, expected) {
  if (operator === 'equals') return cell === expected;
  if (operator === 'not_equals') return cell !== expected;
  if (operator === 'contains') return cell.includes(expected);
  if (operator === 'not_contains') return !cell.includes(expected);
  if (operator === 'starts_with') return cell.startsWith(expected);
  if (operator === 'ends_with') return cell.endsWith(expected);
  throw new Error('Choose a filter operator.');
}

function numberMatches(cell, operator, expected) {
  const parsed = numericText(cell);
  if (!parsed) return false;
  const value = Object.is(parsed.value, -0) ? 0 : parsed.value;
  if (operator === 'gt') return value > expected;
  if (operator === 'gte') return value >= expected;
  if (operator === 'lt') return value < expected;
  if (operator === 'lte') return value <= expected;
  throw new Error('Choose a filter operator.');
}

function compileCondition(condition, width) {
  const operator = condition && condition.operator;
  if (!OPERATOR_IDS.has(operator)) throw new Error('Choose a filter operator.');
  const index = columnIndex(condition, width);
  const spec = OPERATORS.find((item) => item.id === operator);
  if (spec.kind === 'empty') {
    return (row) => {
      const cell = cellString(row[index]);
      return operator === 'empty' ? cell === '' : cell !== '';
    };
  }
  const expectedText = comparisonValue(condition);
  if (expectedText === '') {
    throw new Error('Enter a comparison value, or use “is empty” for a blank cell.');
  }
  if (spec.kind === 'number') {
    const expected = plainNumber(expectedText, 'The comparison value');
    return (row) => numberMatches(cellString(row[index]), operator, expected);
  }
  return (row) => textMatches(cellString(row[index]), operator, expectedText);
}

export function filterRows(table, conditions, mode) {
  const width = requireTable(table);
  if (mode !== 'and' && mode !== 'or') {
    throw new Error('Choose AND or OR. This page does not mix both in one filter.');
  }
  if (!Array.isArray(conditions) || conditions.length === 0) {
    throw new Error('Add at least one condition. This page does not keep every row when no condition is set.');
  }
  const tests = conditions.map((condition) => compileCondition(condition, width));
  const rows = [];
  for (const row of table.rows) {
    if (!Array.isArray(row)) throw new Error('A data row is not a list of cells.');
    const matched = tests.map((test) => test(row));
    const keep = mode === 'and' ? matched.every(Boolean) : matched.some(Boolean);
    if (!keep) continue;
    const cells = [];
    for (let index = 0; index < width; index += 1) cells.push(cellString(row[index]));
    rows.push(cells);
  }
  return {
    headers: table.headers.slice(),
    rows,
    matchCount: rows.length,
    dataRowCount: table.rows.length,
    mode,
  };
}

export function previewMatches(result, limit = PREVIEW_ROW_LIMIT) {
  if (!result || !Array.isArray(result.rows)) throw new Error('There are no filter results to preview.');
  const count = Number(limit);
  if (!Number.isInteger(count) || count < 0) throw new Error('The preview limit is not valid.');
  return result.rows.slice(0, count);
}
