import {
  EXAMPLE_CSV,
  LIMITS,
  PREVIEW_GROUP_LIMIT,
  aggregateTwoRowPivot,
  openXlsx,
  parseCsv,
  previewResult,
  serializeCsv,
} from './pivot-two-row-dimensions.js';

const fileInput = document.querySelector('#file');
const delimiterInput = document.querySelector('#delimiter');
const sheetWrap = document.querySelector('#sheet-wrap');
const sheetInput = document.querySelector('#sheet');
const sourceInput = document.querySelector('#source');
const row1Input = document.querySelector('#row1-field');
const row2Input = document.querySelector('#row2-field');
const columnInput = document.querySelector('#column-field');
const valueInput = document.querySelector('#value-field');
const loadButton = document.querySelector('#load');
const exampleButton = document.querySelector('#example');
const runButton = document.querySelector('#run');
const downloadButton = document.querySelector('#download');
const statusNode = document.querySelector('#tool-status');
const preview = document.querySelector('#preview');

let table = null;
let workbook = null;
let result = null;

function showError(error) {
  table = null;
  result = null;
  for (const select of [row1Input, row2Input, columnInput, valueInput]) {
    select.replaceChildren();
    select.disabled = true;
  }
  runButton.disabled = true;
  downloadButton.disabled = true;
  preview.hidden = true;
  preview.replaceChildren();
  preview.dataset.ready = '0';
  statusNode.textContent = error instanceof Error ? error.message : 'The file could not be calculated.';
  statusNode.classList.add('error');
}

function setStatus(message) {
  statusNode.textContent = message;
  statusNode.classList.remove('error');
}

function addOption(select, value, label) {
  const option = document.createElement('option');
  option.value = value;
  option.textContent = label;
  select.appendChild(option);
}

function fillSelect(select, headers, preferred, blankLabel) {
  select.replaceChildren();
  if (blankLabel) addOption(select, '', blankLabel);
  headers.forEach((name, index) => addOption(select, String(index), name));
  select.disabled = false;
  select.value = preferred == null || preferred === '' ? '' : String(preferred);
}

function hideSheets() {
  workbook = null;
  sheetWrap.hidden = true;
  sheetInput.replaceChildren();
  sheetInput.disabled = true;
}

function showSheets(names) {
  sheetInput.replaceChildren();
  names.forEach((name, index) => addOption(sheetInput, String(index), name));
  sheetInput.disabled = false;
  sheetInput.value = '0';
  sheetWrap.hidden = false;
}

function renderResult(calculated) {
  result = calculated;
  const visible = previewResult(calculated, PREVIEW_GROUP_LIMIT);
  preview.replaceChildren();
  const tableNode = document.createElement('table');
  const caption = document.createElement('caption');
  const row1Name = table.headers[Number(row1Input.value)] || 'Row 1';
  const row2Name = table.headers[Number(row2Input.value)] || 'Row 2';
  const columnName = columnInput.value === '' ? '' : table.headers[Number(columnInput.value)];
  caption.textContent = columnName
    ? `${row1Name} by ${row2Name}, split by ${columnName}`
    : `${row1Name} by ${row2Name}`;
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
  downloadButton.disabled = false;
  runButton.disabled = false;
  const rowNoun = calculated.dataRowCount === 1 ? 'row' : 'rows';
  const sheetName = table && table.name ? `Worksheet ${table.name}: ` : '';
  const previewNote = visible.length < calculated.rows.length
    ? ` Preview shows the first ${PREVIEW_GROUP_LIMIT} groups plus the all-rows row. Download CSV for every group.`
    : ' Download CSV includes every group.';
  const totalNote = calculated.numericCount === 0
    ? 'No plain number was found, so Sum and Average are blank.'
    : `All-rows sum is ${calculated.sum}. All-rows average is ${calculated.average}.`;
  setStatus(`${sheetName}${calculated.detailCount} group ${calculated.detailCount === 1 ? 'row' : 'rows'} from ${calculated.dataRowCount} data ${rowNoun}. ${totalNote}${previewNote}`);
}

function clearResult(message) {
  result = null;
  downloadButton.disabled = true;
  preview.hidden = true;
  preview.replaceChildren();
  preview.dataset.ready = '0';
  statusNode.textContent = message;
  statusNode.classList.add('error');
}

function runCalc() {
  if (!table) {
    showError(new Error('Choose a CSV or XLSX file, or paste CSV, before calculating.'));
    return;
  }
  try {
    const calculated = aggregateTwoRowPivot(
      table,
      row1Input.value,
      row2Input.value,
      valueInput.value,
      columnInput.value,
    );
    renderResult(calculated);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'The file could not be calculated.';
    clearResult(message);
  }
}

async function useTable(nextTable, nextWorkbook) {
  table = nextTable;
  workbook = nextWorkbook;
  const last = table.headers.length - 1;
  fillSelect(row1Input, table.headers, 0, '');
  fillSelect(
    row2Input,
    table.headers,
    table.headers.length > 1 ? 1 : '',
    'Choose a second row field',
  );
  fillSelect(columnInput, table.headers, '', 'No column field');
  fillSelect(valueInput, table.headers, last, '');
  runCalc();
}

function assertFileSize(file) {
  if (file.size > LIMITS.maxBytes) throw new Error('This file is larger than 10 MB.');
}

async function loadFromCurrent() {
  result = null;
  downloadButton.disabled = true;
  preview.hidden = true;
  preview.dataset.ready = '0';
  const file = fileInput.files && fileInput.files[0];
  try {
    if (file) {
      assertFileSize(file);
      if (/\.xlsx$/i.test(file.name)) {
        const opened = await openXlsx(await file.arrayBuffer());
        showSheets(opened.sheetNames);
        const next = await opened.sheet(0);
        workbook = opened;
        await useTable(next, opened);
        return;
      }
      if (/\.xls$/i.test(file.name)) {
        throw new Error('Older .xls workbooks are not supported. Save as .xlsx or CSV.');
      }
      if (!/\.(csv|tsv|txt)$/i.test(file.name)) {
        throw new Error('Choose a CSV or XLSX file.');
      }
      hideSheets();
      const parsed = parseCsv(await file.text(), delimiterInput.value);
      await useTable(parsed, null);
      return;
    }
    hideSheets();
    const parsed = parseCsv(sourceInput.value, delimiterInput.value);
    await useTable(parsed, null);
  } catch (error) {
    hideSheets();
    showError(error);
  }
}

loadButton.addEventListener('click', () => {
  loadFromCurrent();
});

exampleButton.addEventListener('click', () => {
  fileInput.value = '';
  delimiterInput.value = ',';
  sourceInput.value = EXAMPLE_CSV;
  loadFromCurrent();
});

runButton.addEventListener('click', () => {
  runCalc();
});

fileInput.addEventListener('change', () => {
  if (fileInput.files && fileInput.files[0]) loadFromCurrent();
});

sheetInput.addEventListener('change', async () => {
  if (!workbook) return;
  try {
    const next = await workbook.sheet(Number(sheetInput.value));
    await useTable(next, workbook);
  } catch (error) {
    showError(error);
  }
});

downloadButton.addEventListener('click', () => {
  if (!result) return;
  const blob = new Blob([serializeCsv(result.headers, result.rows)], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'pivot-two-row-dimensions.csv';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

row1Input.disabled = true;
row2Input.disabled = true;
columnInput.disabled = true;
valueInput.disabled = true;
runButton.disabled = true;
downloadButton.disabled = true;
