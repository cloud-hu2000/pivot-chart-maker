/**
 * Grouped percent of total, and difference versus the previous period.
 *
 * Group identity is the exact cell text. Null and missing cells are the empty
 * string. Spaces and letter case are kept, so "East" and "east" are different
 * groups.
 *
 * Period identity is the trimmed cell text. "2024-01" and " 2024-01 " are one
 * period. "2024-01" and "2024-1" stay two periods even when they fall on the
 * same instant.
 *
 * A trimmed period is a date when it is a real year (YYYY), year-month
 * (YYYY-M or YYYY-MM, hyphen or slash), year-month-day (one or two digit month
 * and day, hyphen or slash), or an ISO date-time
 * (YYYY-MM-DDTHH:mm:ss with an optional fraction and an optional Z or numeric
 * offset). A date-time with no offset is UTC. The first three fractional
 * digits are milliseconds; further digits are ignored. Invalid calendar dates,
 * including 2024-02-31 and 2023-02-29, are not dates. Excel serial numbers are
 * not turned into calendar dates.
 *
 * Date periods sort by UTC instant, oldest first, and before every non-date.
 * Labels that share an instant stay separate and then sort by the label text.
 * Other periods sort lexicographically after normalization: a year-month or
 * year-month-day shape is zero-padded for the sort key only. The label is not
 * rewritten. Remaining text sorts by UTF-16 code units, so Q1, Q10, Q2 sort
 * as Q1, Q10, Q2.
 *
 * The previous period of P is the immediately previous label in that one
 * sorted list of unique periods from the whole file. A calendar month that
 * does not appear is not inserted. If a group has no row in that previous
 * period, or that row has no plain number, Difference and Percent change are
 * n/a. They are not filled with zero.
 *
 * The measure is a finite plain number (optional sign, decimal, or exponent).
 * Text, blanks, currency, percent signs, and thousands separators are ignored,
 * not treated as zero. Negative zero is zero. A group-period with no plain
 * number has a blank Sum. The grand total is the sum of those group-period
 * sums. Percent of total is Sum / grand total × 100. A missing or zero grand
 * total makes every percent n/a. Percent change is (current − previous) /
 * previous × 100, and is n/a when the previous sum is zero or missing. The
 * page does not show NaN or Infinity.
 *
 * Integer totals use exact arithmetic through 9,007,199,254,740,991. A larger
 * finite integer, or a total past that limit, stops with an error. Percents
 * round half away from zero to 6 decimal places, then drop trailing zeros.
 *
 * Files are read locally by the caller. This module does not upload them.
 */

import { numericText } from './pivot-min-max-distinct.js';

export {
  LIMITS,
  assertByteSize,
  openXlsx,
  parseCsv,
  serializeCsv,
} from './pivot-min-max-distinct.js';

export const NA = 'n/a';
export const PERCENT_DECIMALS = 6;
export const PREVIEW_ROW_LIMIT = 12;
export const RESULT_HEADERS = [
  'Group',
  'Period',
  'Sum',
  'Percent of total',
  'Previous period',
  'Difference',
  'Percent change',
];

export const EXAMPLE_CSV = [
  'Region,Period,Amount',
  'East,2024-01,40',
  'East,2024-01,10',
  'East,2024-02,80',
  'East,2024-04,20',
  'West,2024-04,50',
  'West,2024-01,n/a',
  'North,2024-02,0',
].join('\n');

const DATE_ONLY = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/;
const YEAR_MONTH = /^(\d{4})[-/](\d{1,2})$/;
const YEAR_ONLY = /^(\d{4})$/;
const DATE_TIME = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$/;
const PAD_SHAPE = /^(\d{4})[-/](\d{1,2})(?:[-/](\d{1,2}))?$/;

function cellString(value) {
  return value == null ? '' : String(value);
}

function validYMD(year, month, day) {
  if (year < 0 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return false;
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

export function parsePeriod(label) {
  const text = cellString(label).trim();
  let match = DATE_TIME.exec(text);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = Number(match[4]);
    const minute = Number(match[5]);
    const second = Number(match[6]);
    if (!validYMD(year, month, day) || hour > 23 || minute > 59 || second > 59) return null;
    const fraction = match[7] ? match[7].slice(1) : '';
    const millis = fraction ? Number((fraction + '000').slice(0, 3)) : 0;
    let utc = Date.UTC(year, month - 1, day, hour, minute, second, millis);
    const zone = match[8];
    if (zone && zone !== 'Z') {
      const sign = zone[0] === '-' ? -1 : 1;
      const digits = zone.slice(1).replace(':', '');
      const zoneHour = Number(digits.slice(0, 2));
      const zoneMinute = Number(digits.slice(2, 4));
      if (zoneHour > 23 || zoneMinute > 59) return null;
      utc -= sign * (zoneHour * 60 + zoneMinute) * 60 * 1000;
    }
    return utc;
  }
  match = DATE_ONLY.exec(text);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    if (!validYMD(year, month, day)) return null;
    return Date.UTC(year, month - 1, day);
  }
  match = YEAR_MONTH.exec(text);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    if (year > 9999 || month < 1 || month > 12) return null;
    return Date.UTC(year, month - 1, 1);
  }
  match = YEAR_ONLY.exec(text);
  if (match) return Date.UTC(Number(match[1]), 0, 1);
  return null;
}

function lexicalKey(label) {
  const match = PAD_SHAPE.exec(label);
  if (!match) return label;
  const month = match[2].padStart(2, '0');
  if (match[3] == null) return `${match[1]}-${month}`;
  return `${match[1]}-${month}-${match[3].padStart(2, '0')}`;
}

export function comparePeriods(left, right) {
  const leftTime = parsePeriod(left);
  const rightTime = parsePeriod(right);
  if (leftTime != null && rightTime != null) {
    if (leftTime !== rightTime) return leftTime < rightTime ? -1 : 1;
  } else if (leftTime != null) {
    return -1;
  } else if (rightTime != null) {
    return 1;
  } else {
    const leftKey = lexicalKey(left);
    const rightKey = lexicalKey(right);
    if (leftKey < rightKey) return -1;
    if (leftKey > rightKey) return 1;
  }
  if (left < right) return -1;
  if (left > right) return 1;
  return 0;
}

export function sortPeriods(labels) {
  return [...labels].sort(comparePeriods);
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

export function formatDecimal(value, decimals = PERCENT_DECIMALS) {
  if (!Number.isFinite(value)) return NA;
  if (value === 0 || Object.is(value, -0)) return '0';
  const negative = value < 0;
  const abs = Math.abs(value);
  const extra = 6;
  const match = /^(\d)\.(\d+)e([+-]?\d+)$/.exec(abs.toExponential(decimals + extra));
  if (!match) throw new Error('A result could not be formatted exactly.');
  const digits = match[1] + match[2];
  const exponent = Number(match[3]);
  const power = exponent - (digits.length - 1) + decimals;
  let whole;
  if (power >= 0) {
    whole = BigInt(digits) * (10n ** BigInt(power));
  } else {
    const divisor = 10n ** BigInt(-power);
    const coefficient = BigInt(digits);
    const quotient = coefficient / divisor;
    const remainder = coefficient % divisor;
    whole = quotient + (remainder * 2n >= divisor ? 1n : 0n);
  }
  return finishDecimal(whole, decimals, negative);
}

function formatScaledPercent(numerator, denominator) {
  const negative = (numerator < 0n) !== (denominator < 0n);
  const num = numerator < 0n ? -numerator : numerator;
  const den = denominator < 0n ? -denominator : denominator;
  const scale = 10n ** BigInt(PERCENT_DECIMALS);
  const scaled = num * 100n * scale;
  const quotient = scaled / den;
  const remainder = scaled % den;
  const rounded = quotient + (remainder * 2n >= den ? 1n : 0n);
  return finishDecimal(rounded, PERCENT_DECIMALS, negative);
}

export function formatMeasure(value) {
  if (value == null) return '';
  if (!Number.isFinite(value)) return NA;
  if (value === 0 || Object.is(value, -0)) return '0';
  if (Number.isSafeInteger(value)) return String(value);
  return formatDecimal(value, PERCENT_DECIMALS);
}

export function formatRatioPercent(numerator, denominator) {
  if (numerator == null || denominator == null) return NA;
  if (!Number.isFinite(numerator) || !Number.isFinite(denominator) || denominator === 0) return NA;
  if (Number.isSafeInteger(numerator) && Number.isSafeInteger(denominator)) {
    return formatScaledPercent(BigInt(numerator), BigInt(denominator));
  }
  return formatDecimal((numerator / denominator) * 100, PERCENT_DECIMALS);
}

export function measureNumber(raw) {
  const parsed = numericText(raw);
  if (!parsed) return null;
  const value = Object.is(parsed.value, -0) ? 0 : parsed.value;
  if (!Number.isFinite(value)) return null;
  if (Math.abs(value) > Number.MAX_SAFE_INTEGER || (Number.isInteger(value) && !Number.isSafeInteger(value))) {
    throw new Error('A measure is outside the range that can be totaled exactly.');
  }
  return value;
}

function addMeasure(current, value) {
  const next = current + value;
  if (!Number.isFinite(next) || Math.abs(next) > Number.MAX_SAFE_INTEGER) {
    throw new Error('The measure total is too large to calculate exactly.');
  }
  return Object.is(next, -0) ? 0 : next;
}

function requireIndex(index, width, message) {
  if (!Number.isInteger(index) || index < 0 || index >= width) throw new Error(message);
}

export function aggregatePercentChange(table, groupIndex, periodIndex, measureIndex) {
  if (!table || !Array.isArray(table.headers) || table.headers.length === 0) {
    throw new Error('A header row is required.');
  }
  if (!Array.isArray(table.rows)) throw new Error('Data rows are missing.');
  const width = table.headers.length;
  requireIndex(groupIndex, width, 'Choose a group field from the header.');
  requireIndex(periodIndex, width, 'Choose a period field from the header.');
  requireIndex(measureIndex, width, 'Choose a measure field from the header.');
  if (groupIndex === periodIndex) {
    throw new Error('Choose a period field that is different from the group field.');
  }
  if (table.rows.length === 0) throw new Error('The file has a header but no data rows.');

  const groups = new Map();
  const periodSet = new Set();
  for (const row of table.rows) {
    if (!Array.isArray(row)) throw new Error('A data row is not a list of cells.');
    const group = cellString(row[groupIndex]);
    const period = cellString(row[periodIndex]).trim();
    const amount = measureNumber(cellString(row[measureIndex]));
    periodSet.add(period);
    let bucket = groups.get(group);
    if (!bucket) {
      bucket = new Map();
      groups.set(group, bucket);
    }
    let cell = bucket.get(period);
    if (!cell) {
      cell = { sum: 0, numeric: false };
      bucket.set(period, cell);
    }
    if (amount != null) {
      cell.sum = cell.numeric ? addMeasure(cell.sum, amount) : amount;
      cell.numeric = true;
    }
  }

  const periods = sortPeriods([...periodSet]);
  const previousOf = new Map(periods.map((period, index) => [period, index === 0 ? null : periods[index - 1]]));
  let grand = null;
  for (const bucket of groups.values()) {
    for (const cell of bucket.values()) {
      if (!cell.numeric) continue;
      grand = grand == null ? cell.sum : addMeasure(grand, cell.sum);
    }
  }

  const rows = [];
  for (const [group, bucket] of groups) {
    for (const period of sortPeriods([...bucket.keys()])) {
      const cell = bucket.get(period);
      const sum = cell.numeric ? cell.sum : null;
      const previous = previousOf.get(period);
      let difference = NA;
      let percentChange = NA;
      if (previous != null && bucket.has(previous) && bucket.get(previous).numeric && sum != null) {
        const previousSum = bucket.get(previous).sum;
        difference = formatMeasure(addMeasure(sum, -previousSum));
        percentChange = formatRatioPercent(sum - previousSum, previousSum);
      }
      rows.push([
        group,
        period,
        formatMeasure(sum),
        formatRatioPercent(sum, grand),
        previous == null ? '' : previous,
        difference,
        percentChange,
      ]);
    }
  }

  return {
    headers: RESULT_HEADERS.slice(),
    rows,
    groupCount: groups.size,
    periodCount: periods.length,
    dataRowCount: table.rows.length,
    periods,
    grandTotal: grand == null ? null : formatMeasure(grand),
    grandTotalMissing: grand == null,
    grandTotalZero: grand === 0,
  };
}

export function previewResult(result, limit = PREVIEW_ROW_LIMIT) {
  if (result.rows.length <= limit) return result.rows.slice();
  return result.rows.slice(0, limit);
}
