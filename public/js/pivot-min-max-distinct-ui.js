import {
  EXAMPLE_CSV,
  LIMITS,
  PREVIEW_GROUP_LIMIT,
  aggregateMinMaxDistinct,
  openXlsx,
  parseCsv,
  previewResult,
  serializeCsv,
} from './pivot-min-max-distinct.js';

const fileInput = document.querySelector('#file');
const delimiterInput = document.querySelector('#delimiter');
const sheetWrap = document.querySelector('#sheet-wrap');
const sheetInput = document.querySelector('#sheet');
const sourceInput = document.querySelector('#source');
const groupInput = document.querySelector('#group-field');
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
  groupInput.replaceChildren();
  valueInput.replaceChildren();
  groupInput.disabled = true;
  valueInput.disabled = true;
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

function fillSelect(select, headers, preferred) {
  select.replaceChildren();
  headers.forEach((name, index) => {
    const option = document.createElement('option');
    option.value = String(index);
    option.textContent = name;
    select.appendChild(option);
  });
  select.disabled = false;
  const choice = Math.min(Math.max(preferred, 0), headers.length - 1);
  select.value = String(choice);
}

function hideSheets() {
  workbook = null;
  sheetWrap.hidden = true;
  sheetInput.replaceChildren();
  sheetInput.disabled = true;
}

function showSheets(names) {
  sheetInput.replaceChildren();
  names.forEach((name, index) => {
    const option = document.createElement('option');
    option.value = String(index);
    option.textContent = name;
    sheetInput.appendChild(option);
  });
  sheetInput.disabled = false;
  sheetInput.value = '0';
  sheetWrap.hidden = false;
}

function renderResult(calculated) {
  result = calculated;
  const visible = previewResult(calculated, PREVIEW_GROUP_LIMIT);
  const shownGroups = visible.filter((row) => row[4] === 'group').length;
  preview.replaceChildren();
  const tableNode = document.createElement('table');
  const caption = document.createElement('caption');
  caption.textContent = 'Grouped minimum, maximum, and distinct count';
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
  const noun = calculated.groupCount === 1 ? 'group' : 'groups';
  const rowNoun = calculated.dataRowCount === 1 ? 'row' : 'rows';
  const sheetName = table && table.name ? `Worksheet ${table.name}: ` : '';
  if (shownGroups < calculated.groupCount) {
    setStatus(`${sheetName}${calculated.groupCount} ${noun} from ${calculated.dataRowCount} data ${rowNoun}. Preview shows the first ${shownGroups} groups. Download CSV for every group and the all-rows summary.`);
  } else {
    setStatus(`${sheetName}${calculated.groupCount} ${noun} from ${calculated.dataRowCount} data ${rowNoun}. Download CSV includes every group and the all-rows summary.`);
  }
}

function runCalc() {
  if (!table) {
    showError(new Error('Choose a CSV or XLSX file, or paste CSV, before calculating.'));
    return;
  }
  try {
    const calculated = aggregateMinMaxDistinct(table, Number(groupInput.value), Number(valueInput.value));
    renderResult(calculated);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'The file could not be calculated.';
    result = null;
    downloadButton.disabled = true;
    preview.hidden = true;
    preview.replaceChildren();
    preview.dataset.ready = '0';
    statusNode.textContent = message;
    statusNode.classList.add('error');
  }
}

async function useTable(nextTable, nextWorkbook) {
  table = nextTable;
  workbook = nextWorkbook;
  fillSelect(groupInput, table.headers, 0);
  fillSelect(valueInput, table.headers, table.headers.length > 1 ? 1 : 0);
  runCalc();
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

function assertFileSize(file) {
  if (file.size > LIMITS.maxBytes) throw new Error('This file is larger than 10 MB.');
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
  link.download = 'pivot-min-max-distinct.csv';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

groupInput.disabled = true;
valueInput.disabled = true;
runButton.disabled = true;
downloadButton.disabled = true;
