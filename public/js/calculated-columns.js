/**
 * Append one calculated column by arithmetic on the current row.
 *
 * The formula may name existing headers in square brackets and may use
 * numbers, parentheses, and the operators + - * /. * and / run before + and
 * -. Operators at the same level run from left to right. A leading + or -
 * is allowed. The formula must name at least one column. Functions, ranges,
 * comparisons, and references to other rows are rejected.
 *
 * A referenced cell is used only when numericText accepts it. Blank cells,
 * words, currency symbols, percent signs, and thousands separators are not
 * numbers and are not treated as zero. If any referenced cell on a row is
 * not a plain number, the new cell is blank. Division by zero and a result
 * that is not a finite JavaScript number also leave that cell blank. They
 * are not written as Infinity or NaN. Negative zero is used as zero.
 *
 * The new column is appended. Existing cells are copied unchanged. The
 * result is static text, not a spreadsheet formula. Files are read locally
 * by the caller.
 */

import {
  LIMITS,
  assertByteSize,
  numericText,
  openXlsx,
  parseCsv,
  serializeCsv,
} from './pivot-min-max-distinct.js';

export {
  LIMITS,
  assertByteSize,
  openXlsx,
  parseCsv,
  serializeCsv,
};

export const PREVIEW_ROW_LIMIT = 12;
export const MAX_FORMULA_LENGTH = 2000;
export const MAX_COLUMN_NAME_LENGTH = 200;
export const MAX_PAREN_DEPTH = 32;

export const EXAMPLE_CSV = [
  'Item,Width,Height',
  'tile,2,3',
  'panel,10,4',
  'label,n/a,1',
  'spacer,,5',
  'strip,1.5,2',
].join('\n');

export const EXAMPLE_FORMULA = '[Width] + [Height] * 2';
export const EXAMPLE_COLUMN_NAME = 'Calculated';

const NUMBER_LITERAL = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;

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

function normalizeNewName(columnName) {
  if (columnName == null) throw new Error('Enter a name for the new column.');
  const name = String(columnName).trim();
  if (name === '') throw new Error('Enter a name for the new column.');
  if (name.length > MAX_COLUMN_NAME_LENGTH) {
    throw new Error('The new column name is longer than 200 characters.');
  }
  return name;
}

function unexpected(formula, index) {
  const ch = formula[index];
  throw new Error(`The formula has an unexpected ${JSON.stringify(ch)} at position ${index + 1}.`);
}

function tokenize(formula) {
  const tokens = [];
  let index = 0;
  while (index < formula.length) {
    const ch = formula[index];
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      index += 1;
      continue;
    }
    if ('+-*/()'.includes(ch)) {
      tokens.push({ type: 'op', value: ch });
      index += 1;
      continue;
    }
    if (ch === '[') {
      const end = formula.indexOf(']', index + 1);
      if (end < 0) throw new Error('A column name is missing its closing ].');
      const name = formula.slice(index + 1, end);
      if (name === '') throw new Error('A column name inside [] is empty.');
      tokens.push({ type: 'col', name });
      index = end + 1;
      continue;
    }
    if (ch === ']') throw new Error('A ] does not have a matching [.');
    if (ch === '.' || (ch >= '0' && ch <= '9')) {
      const match = NUMBER_LITERAL.exec(formula.slice(index));
      if (!match) unexpected(formula, index);
      const text = match[0];
      const value = Number(text);
      if (!Number.isFinite(value)) {
        throw new Error(`The number ${text} is too large to use in the formula.`);
      }
      tokens.push({ type: 'num', value: Object.is(value, -0) ? 0 : value });
      index += text.length;
      continue;
    }
    unexpected(formula, index);
  }
  return tokens;
}

function parseFormula(formula) {
  if (formula == null || String(formula).trim() === '') {
    throw new Error('Enter a formula that names an existing column, such as [Width] + [Height].');
  }
  const source = String(formula);
  if (source.length > MAX_FORMULA_LENGTH) {
    throw new Error('The formula is longer than 2,000 characters.');
  }
  const tokens = tokenize(source);
  if (tokens.length === 0) {
    throw new Error('Enter a formula that names an existing column, such as [Width] + [Height].');
  }
  let pos = 0;
  let depth = 0;
  function peek() {
    return tokens[pos];
  }
  function eat() {
    const token = tokens[pos];
    pos += 1;
    return token;
  }
  function parseExpr() {
    let node = parseTerm();
    while (peek() && peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
      const op = eat().value;
      node = { type: 'bin', op, left: node, right: parseTerm() };
    }
    return node;
  }
  function parseTerm() {
    let node = parseUnary();
    while (peek() && peek().type === 'op' && (peek().value === '*' || peek().value === '/')) {
      const op = eat().value;
      node = { type: 'bin', op, left: node, right: parseUnary() };
    }
    return node;
  }
  function parseUnary() {
    if (peek() && peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
      const op = eat().value;
      return { type: 'unary', op, arg: parseUnary() };
    }
    return parsePrimary();
  }
  function parsePrimary() {
    const token = peek();
    if (!token) throw new Error('The formula ended before a value.');
    if (token.type === 'num' || token.type === 'col') {
      eat();
      return token.type === 'num'
        ? { type: 'num', value: token.value }
        : { type: 'col', name: token.name };
    }
    if (token.type === 'op' && token.value === '(') {
      eat();
      depth += 1;
      if (depth > MAX_PAREN_DEPTH) {
        throw new Error('This formula has too many nested parentheses.');
      }
      if (peek() && peek().type === 'op' && peek().value === ')') {
        throw new Error('The parentheses are empty.');
      }
      const inner = parseExpr();
      if (!peek() || peek().type !== 'op' || peek().value !== ')') {
        throw new Error('The formula is missing a closing parenthesis.');
      }
      eat();
      depth -= 1;
      return inner;
    }
    if (token.type === 'op' && token.value === ')') {
      throw new Error('The formula has an extra closing parenthesis.');
    }
    throw new Error('The formula is missing a value before an operator.');
  }
  const ast = parseExpr();
  if (peek()) {
    if (peek().type === 'op' && peek().value === ')') {
      throw new Error('The formula has an extra closing parenthesis.');
    }
    throw new Error('Enter an operator between values. This page does not multiply by writing values next to each other.');
  }
  return ast;
}

function collectRefs(node, out) {
  if (node.type === 'col') out.push(node.name);
  else if (node.type === 'unary') collectRefs(node.arg, out);
  else if (node.type === 'bin') {
    collectRefs(node.left, out);
    collectRefs(node.right, out);
  }
  return out;
}

function finish(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return { kind: 'not_finite' };
  return Object.is(value, -0) ? 0 : value;
}

function evaluate(node, values) {
  if (node.type === 'num') return node.value;
  if (node.type === 'col') {
    const value = values.get(node.name);
    if (typeof value !== 'number') throw new Error('The formula could not be calculated.');
    return value;
  }
  if (node.type === 'unary') {
    const value = evaluate(node.arg, values);
    if (typeof value !== 'number') return value;
    return finish(node.op === '-' ? -value : value);
  }
  const left = evaluate(node.left, values);
  if (typeof left !== 'number') return left;
  const right = evaluate(node.right, values);
  if (typeof right !== 'number') return right;
  if (node.op === '/' && right === 0) return { kind: 'divide_by_zero' };
  if (node.op === '+') return finish(left + right);
  if (node.op === '-') return finish(left - right);
  if (node.op === '*') return finish(left * right);
  if (node.op === '/') return finish(left / right);
  throw new Error('The formula could not be calculated.');
}

function formatCalculated(value) {
  return String(Object.is(value, -0) ? 0 : value);
}

export function calculateColumn(table, formula, columnName) {
  const width = requireTable(table);
  if (table.rows.length > LIMITS.maxDataRows) {
    throw new Error('This file has more than 100,000 data rows.');
  }
  if (width >= LIMITS.maxColumns) {
    throw new Error('This file already has 1,000 columns. This page cannot add another.');
  }
  const name = normalizeNewName(columnName);
  if (table.headers.includes(name)) {
    throw new Error('That column name is already in the file. This page adds a new column and does not replace an existing one.');
  }
  const ast = parseFormula(formula);
  const refs = [...new Set(collectRefs(ast, []))];
  if (refs.length === 0) {
    throw new Error('The formula must name at least one existing column in square brackets, such as [Width].');
  }
  const headerIndex = new Map(table.headers.map((header, index) => [header, index]));
  for (const ref of refs) {
    if (!headerIndex.has(ref)) {
      throw new Error(`There is no column named ${JSON.stringify(ref)}. Column names must match exactly, including spaces and letter case.`);
    }
  }
  const rows = [];
  let calculatedCount = 0;
  let blankNotPlain = 0;
  let blankDivideByZero = 0;
  let blankNotFinite = 0;
  for (const source of table.rows) {
    if (!Array.isArray(source)) throw new Error('A data row is not a list of cells.');
    const cells = [];
    for (let index = 0; index < width; index += 1) cells.push(cellString(source[index]));
    let usable = true;
    const values = new Map();
    for (const ref of refs) {
      const parsed = numericText(cells[headerIndex.get(ref)]);
      if (!parsed) {
        usable = false;
        break;
      }
      values.set(ref, Object.is(parsed.value, -0) ? 0 : parsed.value);
    }
    let calculated = '';
    if (!usable) {
      blankNotPlain += 1;
    } else {
      const result = evaluate(ast, values);
      if (typeof result !== 'number') {
        if (result.kind === 'divide_by_zero') blankDivideByZero += 1;
        else blankNotFinite += 1;
      } else {
        calculated = formatCalculated(result);
        calculatedCount += 1;
      }
    }
    cells.push(calculated);
    rows.push(cells);
  }
  return {
    headers: table.headers.concat(name),
    rows,
    dataRowCount: table.rows.length,
    calculatedCount,
    blankNotPlain,
    blankDivideByZero,
    blankNotFinite,
    columnName: name,
  };
}

export function previewRows(result, limit = PREVIEW_ROW_LIMIT) {
  if (!result || !Array.isArray(result.rows)) throw new Error('There is no calculated column to preview.');
  const count = Number(limit);
  if (!Number.isInteger(count) || count < 0) throw new Error('The preview limit is not valid.');
  return result.rows.slice(0, count);
}
