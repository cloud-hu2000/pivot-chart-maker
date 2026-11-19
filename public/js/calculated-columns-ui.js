import {
  EXAMPLE_COLUMN_NAME,
  EXAMPLE_CSV,
  EXAMPLE_FORMULA,
  LIMITS,
  PREVIEW_ROW_LIMIT,
  calculateColumn,
  openXlsx,
  parseCsv,
  previewRows,
  serializeCsv,
} from './calculated-columns.js';

const fileInput = document.querySelector('#file');
const delimiterInput = document.querySelector('#delimiter');
const sheetWrap = document.querySelector('#sheet-wrap');
const sheetInput = document.querySelector('#sheet');
const sourceInput = document.querySelector('#source');
const formulaInput = document.querySelector('#formula');
const columnNameInput = document.querySelector('#column-name');
const columnList = document.querySelector('#column-names');
const nameNote = document.querySelector('#name-note');
const loadButton = document.querySelector('#load');
const exampleButton = document.querySelector('#example');
const runButton = document.querySelector('#run');
const downloadButton = document.querySelector('#download');
const statusNode = document.querySelector('#tool-status');
const preview = document.querySelector('#preview');

let table = null;
let workbook = null;
let result = null;

function setStatus(message, isError) {
  statusNode.textContent = message;
  statusNode.classList.toggle('error', Boolean(isError));
}

function clearPreview() {
  result = null;
  downloadButton.disabled = true;
  preview.hidden = true;
  preview.replaceChildren();
  preview.dataset.ready = '0';
  delete preview.dataset.rowCount;
  delete preview.dataset.calculatedCount;
  delete preview.dataset.blankNotPlain;
}

function hideSheets() {
  workbook = null;
  sheetWrap.hidden = true;
  sheetInput.replaceChildren();
  sheetInput.disabled = true;
}

function showLoadError(error) {
  table = null;
  hideSheets();
  clearPreview();
  renderColumnNames();
  setStatus(error instanceof Error ? error.message : 'The file could not be read.', true);
}

function showRunError(error) {
  clearPreview();
  setStatus(error instanceof Error ? error.message : 'The column could not be calculated.', true);
}

function addOption(select, value, label) {
  const option = document.createElement('option');
  option.value = value;
  option.textContent = label;
  select.appendChild(option);
}

function renderColumnNames() {
  columnList.replaceChildren();
  const headers = table ? table.headers : [];
  if (!headers.length) {
    const item = document.createElement('li');
    item.textContent = 'Load data to see column names.';
    columnList.appendChild(item);
    nameNote.textContent = '';
    return;
  }
  for (const header of headers) {
    const item = document.createElement('li');
    item.textContent = header;
    columnList.appendChild(item);
  }
  const blocked = headers.filter((header) => header.includes(']'));
  nameNote.textContent = blocked.length
    ? `The column name ${blocked.map((header) => JSON.stringify(header)).join(', ')} contains ]. This page cannot name it in a formula.`
    : '';
}

function rowSentence(count, text) {
  if (count === 0) return '';
  const noun = count === 1 ? 'row is' : 'rows are';
  return `${count} ${noun} blank because ${text}.`;
}

function renderResult(calculated) {
  result = calculated;
  const visible = previewRows(calculated, PREVIEW_ROW_LIMIT);
  preview.replaceChildren();
  const tableNode = document.createElement('table');
  const caption = document.createElement('caption');
  caption.textContent = `Original columns plus ${calculated.columnName}`;
  tableNode.appendChild(caption);
  const head = document.createElement('thead');
  const headRow = document.createElement('tr');
  for (const header of calculated.headers) {
    const cell = document.createElement('th');
    cell.scope = 'col';
    cell.textContent = header;
    headRow.appendChild(cell);
  }
  head.appendChild(headRow);
  tableNode.appendChild(head);
  const body = document.createElement('tbody');
  for (const row of visible) {
    const tr = document.createElement('tr');
    for (const value of row) {
      const cell = document.createElement('td');
      cell.textContent = value;
      tr.appendChild(cell);
    }
    body.appendChild(tr);
  }
  tableNode.appendChild(body);
  preview.appendChild(tableNode);
  preview.hidden = false;
  preview.dataset.ready = '1';
  preview.dataset.rowCount = String(calculated.dataRowCount);
  preview.dataset.calculatedCount = String(calculated.calculatedCount);
  preview.dataset.blankNotPlain = String(calculated.blankNotPlain);
  downloadButton.disabled = false;
  const sheetName = table && table.name ? `Worksheet ${table.name}: ` : '';
  const blanks = [
    rowSentence(calculated.blankNotPlain, 'a referenced cell is not a plain number'),
    rowSentence(calculated.blankDivideByZero, 'a divisor is zero'),
    rowSentence(calculated.blankNotFinite, 'the result is not a finite number'),
  ].filter(Boolean).join(' ');
  const blankText = blanks || 'Every calculated cell is a finite number.';
  const previewText = calculated.rows.length > visible.length
    ? `The preview shows the first ${PREVIEW_ROW_LIMIT} rows. The download contains all ${calculated.dataRowCount} rows.`
    : 'The preview lists every row. The download contains the original columns and the new column.';
  setStatus(`${sheetName}${calculated.calculatedCount} of ${calculated.dataRowCount} data rows have a number in ${calculated.columnName}. ${blankText} ${previewText}`, false);
}

function runCalculation() {
  if (!table) {
    showRunError(new Error('Choose a CSV or XLSX file, or paste CSV, and click Load data before adding a column.'));
    return;
  }
  try {
    renderResult(calculateColumn(table, formulaInput.value, columnNameInput.value));
  } catch (error) {
    showRunError(error);
  }
}

function markStale(message) {
  if (!result && preview.dataset.ready !== '1') return;
  clearPreview();
  setStatus(message, true);
}

async function useTable(nextTable) {
  table = nextTable;
  clearPreview();
  renderColumnNames();
  const sheetName = table.name ? `Worksheet ${table.name}: ` : '';
  const noun = table.rows.length === 1 ? 'row' : 'rows';
  setStatus(`${sheetName}Loaded ${table.rows.length} data ${noun}. Enter a formula, then click Add column.`, false);
}

function assertFileSize(file) {
  if (file.size > LIMITS.maxBytes) throw new Error('This file is larger than 10 MB.');
}

async function loadFromCurrent() {
  clearPreview();
  const file = fileInput.files && fileInput.files[0];
  try {
    if (file) {
      assertFileSize(file);
      if (/\.xlsx$/i.test(file.name)) {
        const opened = await openXlsx(await file.arrayBuffer());
        workbook = opened;
        sheetInput.replaceChildren();
        opened.sheetNames.forEach((name, index) => addOption(sheetInput, String(index), name));
        sheetInput.disabled = false;
        sheetInput.value = '0';
        sheetWrap.hidden = false;
        await useTable(await opened.sheet(0));
        return;
      }
      if (/\.xls$/i.test(file.name)) {
        throw new Error('Older .xls workbooks are not supported. Save as .xlsx or CSV.');
      }
      if (!/\.(csv|tsv|txt)$/i.test(file.name)) {
        throw new Error('Choose a CSV or XLSX file.');
      }
      hideSheets();
      await useTable(parseCsv(await file.text(), delimiterInput.value));
      return;
    }
    hideSheets();
    await useTable(parseCsv(sourceInput.value, delimiterInput.value));
  } catch (error) {
    showLoadError(error);
  }
}

loadButton.addEventListener('click', () => {
  loadFromCurrent();
});

exampleButton.addEventListener('click', async () => {
  fileInput.value = '';
  delimiterInput.value = ',';
  sourceInput.value = EXAMPLE_CSV;
  formulaInput.value = EXAMPLE_FORMULA;
  columnNameInput.value = EXAMPLE_COLUMN_NAME;
  await loadFromCurrent();
  if (!table) return;
  runCalculation();
});

runButton.addEventListener('click', () => {
  runCalculation();
});

formulaInput.addEventListener('input', () => {
  markStale('The formula changed. Click Add column to update the calculated values.');
});

columnNameInput.addEventListener('input', () => {
  markStale('The new column name changed. Click Add column to update the calculated values.');
});

delimiterInput.addEventListener('change', () => {
  markStale('Delimiter changed. Click Load data, then Add column.');
});

fileInput.addEventListener('change', () => {
  if (fileInput.files && fileInput.files[0]) loadFromCurrent();
});

sheetInput.addEventListener('change', async () => {
  if (!workbook) return;
  try {
    await useTable(await workbook.sheet(Number(sheetInput.value)));
  } catch (error) {
    showLoadError(error);
  }
});

downloadButton.addEventListener('click', () => {
  if (!result) return;
  const blob = new Blob([serializeCsv(result.headers, result.rows)], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'calculated-columns.csv';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

renderColumnNames();
downloadButton.disabled = true;
