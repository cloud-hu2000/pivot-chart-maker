/**
 * Pivot with two row fields and an optional column field.
 *
 * Aggregations are Sum, Average, and Count only. Row and column labels are
 * the exact cell text. Null and missing cells are the empty string. Spaces
 * and letter case are kept, so "East" and "east" are different groups, and
 * "Desk" and "Desk " are different groups. Duplicate labels in either row
 * field, or in the column field, combine into the same group.
 *
 * The second row field is required. If it is missing, or it is the same
 * column as the first row field, this function throws. It does not invent a
 * second dimension, copy the first field, or insert a label such as "(all)"
 * or "(blank)". An empty cell in a selected second field is a real label,
 * the empty string, not a missing field.
 *
 * The column field is optional. Without it, each detail row is one observed
 * (row 1, row 2) pair. With it, each detail row is one observed
 * (row 1, row 2, column) triple. A combination that does not appear in the
 * file is omitted. It is not filled with zero.
 *
 * Sum, Average, and Count use a cell only when its trimmed text is a finite
 * plain number (optional sign, decimal, or exponent), the same rule as
 * numericText. Text, blanks, currency symbols, percent signs, and thousands
 * separators are ignored, not treated as zero. Count is how many of those
 * plain numbers were included, and it is the average's denominator. Average
 * is Sum / Count. A group with no plain number has a blank Sum, a blank
 * Average, and Count 0. Count 0 is the number of numeric cells, not a
 * substitute for the missing sum.
 *
 * The last row is the whole file. Its average is the file sum divided by the
 * file count, not the average of the group averages. Its first label is
 * "(all rows)" and its scope is "all rows". A source group that is itself
 * named "(all rows)" stays a normal group; the scope column tells them apart.
 *
 * Integer totals inside the safe integer range are exact. A plain number, or
 * a running total, outside 9007199254740991 stops with an error. Other
 * decimals are added as JavaScript numbers and then shown to 6 decimal
 * places, half away from zero. The page does not show NaN or Infinity.
 *
 * Files are read locally by the caller. This module does not upload them.
 */

import { numericText } from './pivot-min-max-distinct.js';
import { formatDecimal } from './pivot-percent-change.js';

export {
  LIMITS,
  assertByteSize,
  openXlsx,
  parseCsv,
  serializeCsv,
} from './pivot-min-max-distinct.js';

export const AVERAGE_DECIMALS = 6;
export const SCOPE_GROUP = 'group';
export const SCOPE_OVERALL = 'all rows';
export const OVERALL_LABEL = '(all rows)';
export const PREVIEW_GROUP_LIMIT = 12;
export const BASE_HEADERS = ['Row 1', 'Row 2', 'Sum', 'Average', 'Count', 'Scope'];
export const COLUMN_HEADERS = ['Row 1', 'Row 2', 'Column', 'Sum', 'Average', 'Count', 'Scope'];

export const EXAMPLE_CSV = [
  'Region,Product,Channel,Amount',
  'East,Desk,Online,10',
  'East,Desk,Online,5',
  'East,Desk,Store,7',
  'East,Chair,Online,8',
  'East,Chair,Store,',
  'West,Desk,Online,4',
  'West,Desk,Store,n/a',
  'West,Chair,Online,6',
  'West,Chair,Online,6',
  'North,Desk,Store,3',
  'east,Desk,Online,1',
].join('\n');

const MAX_SAFE = Number.MAX_SAFE_INTEGER;

function cellString(value) {
  return value == null ? '' : String(value);
}

function emptyBucket() {
  return { sum: 0, count: 0, numeric: false };
}

export function measureNumber(raw) {
  const parsed = numericText(raw);
  if (!parsed) return null;
  const value = Object.is(parsed.value, -0) ? 0 : parsed.value;
  if (!Number.isFinite(value)) return null;
  if (Math.abs(value) > MAX_SAFE || (Number.isInteger(value) && !Number.isSafeInteger(value))) {
    throw new Error('A measure is outside the range that can be totaled exactly.');
  }
  return value;
}

function addMeasure(current, value) {
  const next = current + value;
  if (!Number.isFinite(next) || Math.abs(next) > MAX_SAFE) {
    throw new Error('The measure total is too large to calculate exactly.');
  }
  return Object.is(next, -0) ? 0 : next;
}

function absorb(bucket, amount) {
  if (amount == null) return;
  bucket.sum = bucket.numeric ? addMeasure(bucket.sum, amount) : amount;
  bucket.count += 1;
  bucket.numeric = true;
}

function finishDecimal(scaledInt, decimals, negative) {
  let text = scaledInt.toString();
  if (text.length <= decimals) text = text.padStart(decimals + 1, '0');
  const whole = text.slice(0, text.length - decimals);
  const fraction = text.slice(text.length - decimals).replace(/0+$/, '');
  const body = fraction ? `${whole}.${fraction}` : whole;
  if (body === '0') return '0';
  return negative ? `-${body}` : body;
}

function formatExactRatio(numerator, denominator) {
  const negative = (numerator < 0) !== (denominator < 0);
  const num = BigInt(Math.abs(numerator));
  const den = BigInt(Math.abs(denominator));
  const scale = 10n ** BigInt(AVERAGE_DECIMALS);
  const scaled = num * scale;
  const quotient = scaled / den;
  const remainder = scaled % den;
  const rounded = quotient + (remainder * 2n >= den ? 1n : 0n);
  return finishDecimal(rounded, AVERAGE_DECIMALS, negative && rounded !== 0n);
}

export function formatSum(bucket) {
  if (!bucket.numeric) return '';
  const sum = Object.is(bucket.sum, -0) ? 0 : bucket.sum;
  if (!Number.isFinite(sum)) throw new Error('A result could not be formatted.');
  if (sum === 0) return '0';
  if (Number.isSafeInteger(sum)) return String(sum);
  const text = formatDecimal(sum, AVERAGE_DECIMALS);
  if (text === 'n/a') throw new Error('A result could not be formatted.');
  return text;
}

export function formatAverage(bucket) {
  if (!bucket.numeric || bucket.count === 0) return '';
  const sum = Object.is(bucket.sum, -0) ? 0 : bucket.sum;
  if (!Number.isFinite(sum)) throw new Error('A result could not be formatted.');
  if (Number.isSafeInteger(sum) && Number.isSafeInteger(bucket.count)) {
    return formatExactRatio(sum, bucket.count);
  }
  const text = formatDecimal(sum / bucket.count, AVERAGE_DECIMALS);
  if (text === 'n/a') throw new Error('A result could not be formatted.');
  return text;
}

function requireIndex(index, width, message) {
  if (!Number.isInteger(index) || index < 0 || index >= width) throw new Error(message);
}

function asIndex(raw) {
  if (typeof raw === 'number') return raw;
  if (typeof raw === 'string' && raw !== '' && /^-?\d+$/.test(raw)) return Number(raw);
  return raw;
}

function metricCells(bucket) {
  return [formatSum(bucket), formatAverage(bucket), String(bucket.count), SCOPE_GROUP];
}

export function aggregateTwoRowPivot(table, row1Index, row2Index, valueIndex, columnIndex = null) {
  if (!table || !Array.isArray(table.headers) || table.headers.length === 0) {
    throw new Error('A header row is required.');
  }
  if (!Array.isArray(table.rows)) throw new Error('Data rows are missing.');
  const width = table.headers.length;
  if (row2Index == null || row2Index === '') {
    throw new Error('Choose a second row field. This page does not invent a second row dimension.');
  }
  const row1 = asIndex(row1Index);
  const row2 = asIndex(row2Index);
  const value = asIndex(valueIndex);
  const columnOmitted = columnIndex == null || columnIndex === '';
  const column = columnOmitted ? null : asIndex(columnIndex);
  requireIndex(row1, width, 'Choose the first row field from the header.');
  requireIndex(row2, width, 'Choose a second row field from the header.');
  requireIndex(value, width, 'Choose a value field from the header.');
  if (!columnOmitted) {
    requireIndex(column, width, 'Choose a column field from the header, or leave it empty.');
  }
  if (row1 === row2) {
    throw new Error('Choose a second row field that is different from the first.');
  }
  if (!columnOmitted && (column === row1 || column === row2)) {
    throw new Error('Choose a column field that is different from both row fields, or leave it empty.');
  }
  if (table.rows.length === 0) throw new Error('The file has a header but no data rows.');

  const pairs = new Map();
  const pairOrder = [];
  const overall = emptyBucket();

  for (const row of table.rows) {
    if (!Array.isArray(row)) throw new Error('A data row is not a list of cells.');
    const first = cellString(row[row1]);
    const second = cellString(row[row2]);
    const columnLabel = columnOmitted ? null : cellString(row[column]);
    const amount = measureNumber(cellString(row[value]));
    let inner = pairs.get(first);
    if (!inner) {
      inner = new Map();
      pairs.set(first, inner);
    }
    let pair = inner.get(second);
    if (!pair) {
      pair = new Map();
      inner.set(second, pair);
      pairOrder.push([first, second]);
    }
    const columnKey = columnOmitted ? '' : columnLabel;
    let bucket = pair.get(columnKey);
    if (!bucket) {
      bucket = emptyBucket();
      pair.set(columnKey, bucket);
    }
    absorb(bucket, amount);
    absorb(overall, amount);
  }

  const rows = [];
  for (const [first, second] of pairOrder) {
    const pair = pairs.get(first).get(second);
    for (const [columnLabel, bucket] of pair) {
      const cells = [first, second];
      if (!columnOmitted) cells.push(columnLabel);
      cells.push(...metricCells(bucket));
      rows.push(cells);
    }
  }
  const summary = [OVERALL_LABEL, ''];
  if (!columnOmitted) summary.push('');
  summary.push(formatSum(overall), formatAverage(overall), String(overall.count), SCOPE_OVERALL);
  rows.push(summary);

  return {
    headers: (columnOmitted ? BASE_HEADERS : COLUMN_HEADERS).slice(),
    rows,
    pairCount: pairOrder.length,
    detailCount: rows.length - 1,
    dataRowCount: table.rows.length,
    columnField: !columnOmitted,
    numericCount: overall.count,
    sum: formatSum(overall),
    average: formatAverage(overall),
  };
}

export function previewResult(result, limit = PREVIEW_GROUP_LIMIT) {
  const details = result.rows.slice(0, -1);
  const overall = result.rows[result.rows.length - 1];
  if (details.length <= limit) return result.rows.slice();
  return [...details.slice(0, limit), overall];
}
