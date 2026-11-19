import {
  EXAMPLE_CSV,
  EXAMPLE_ROW_INDEX,
  EXAMPLE_VALUE_INDEX,
  LIMITS,
  PREVIEW_GROUP_LIMIT,
  buildExcelPivotTable,
  formatPivotNumber,
  openXlsx,
  parseCsv,
  previewGroups,
} from './export-excel-pivot-table.js';

const fileInput = document.querySelector('#file');
const delimiterInput = document.querySelector('#delimiter');
const sheetWrap = document.querySelector('#sheet-wrap');
const sheetInput = document.querySelector('#sheet');
const sourceInput = document.querySelector('#source');
const rowInput = document.querySelector('#row-field');
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
  delete preview.dataset.groupCount;
}

function hideSheets() {
  workbook = null;
  sheetWrap.hidden = true;
  sheetInput.replaceChildren();
  sheetInput.disabled = true;
}

function addOption(select, value, label) {
  const option = document.createElement('option');
  option.value = value;
  option.textContent = label;
  select.appendChild(option);
}

function fillSelect(select, headers) {
  const previous = select.value;
  select.replaceChildren();
  if (!headers.length) {
    addOption(select, '', 'Load data first');
    select.disabled = true;
    select.value = '';
    return;
  }
  headers.forEach((header, index) => addOption(select, String(index), header));
  select.disabled = false;
  select.value = headers[Number(previous)] ? previous : '0';
}

function refreshFields() {
  const headers = table ? table.headers : [];
  fillSelect(rowInput, headers);
  fillSelect(valueInput, headers);
}

function showLoadError(error) {
  table = null;
  hideSheets();
  clearPreview();
  refreshFields();
  setStatus(error instanceof Error ? error.message : 'The file could not be read.', true);
}

function showBuildError(error) {
  clearPreview();
  setStatus(error instanceof Error ? error.message : 'The workbook could not be built.', true);
}

function sumText(value) {
  return value == null ? '' : formatPivotNumber(value);
}

function renderResult(next) {
  result = next;
  const shown = previewGroups(next);
  const tableNode = document.createElement('table');
  const caption = document.createElement('caption');
  caption.textContent = shown.hiddenGroupCount > 0
    ? `Pivot sheet preview, first ${shown.groups.length} groups. The workbook includes all ${next.groups.length} groups.`
    : 'Pivot sheet preview. The workbook includes every group shown here.';
  tableNode.appendChild(caption);
  const head = document.createElement('thead');
  const headRow = document.createElement('tr');
  for (const label of [next.rowFieldName, next.dataFieldName, 'Plain numbers', 'Ignored cells']) {
    const cell = document.createElement('th');
    cell.textContent = label;
    headRow.appendChild(cell);
  }
  head.appendChild(headRow);
  tableNode.appendChild(head);
  const body = document.createElement('tbody');
  const lines = [
    ...shown.groups.map((group) => [group.label, sumText(group.sum), String(group.numberCount), String(group.ignoredCount)]),
    ['Grand Total', sumText(next.grandSum), String(next.numberCount), String(next.ignoredCount)],
  ];
  for (const line of lines) {
    const row = document.createElement('tr');
    for (const value of line) {
      const cell = document.createElement('td');
      cell.textContent = value;
      row.appendChild(cell);
    }
    body.appendChild(row);
  }
  tableNode.appendChild(body);
  preview.replaceChildren(tableNode);
  preview.hidden = false;
  preview.dataset.ready = '1';
  preview.dataset.groupCount = String(next.groups.length);
  downloadButton.disabled = false;
  const noun = next.groups.length === 1 ? 'group' : 'groups';
  const numberNoun = next.numberCount === 1 ? 'plain number is' : 'plain numbers are';
  const ignoredNoun = next.ignoredCount === 1 ? 'value cell is' : 'value cells are';
  const hidden = shown.hiddenGroupCount > 0 ? ` Preview lists the first ${PREVIEW_GROUP_LIMIT} groups.` : '';
  setStatus(`${next.groups.length} ${noun}. Grand total of plain numbers is ${sumText(next.grandSum)}. ${next.numberCount} ${numberNoun} included. ${next.ignoredCount} ${ignoredNoun} left out of the sum.${hidden}`, false);
}

function runBuild() {
  if (!table) {
    showBuildError(new Error('Load a CSV or XLSX file before building the workbook.'));
    return;
  }
  try {
    renderResult(buildExcelPivotTable(table, Number(rowInput.value), Number(valueInput.value)));
  } catch (error) {
    showBuildError(error);
  }
}

async function useTable(nextTable) {
  table = nextTable;
  clearPreview();
  refreshFields();
  const sheetName = table.name ? `Worksheet ${table.name}: ` : '';
  const noun = table.rows.length === 1 ? 'row' : 'rows';
  setStatus(`${sheetName}Loaded ${table.rows.length} data ${noun}. Choose a row field and a value column, then click Build PivotTable workbook.`, false);
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
  await loadFromCurrent();
  if (!table) return;
  rowInput.value = String(EXAMPLE_ROW_INDEX);
  valueInput.value = String(EXAMPLE_VALUE_INDEX);
  runBuild();
});

runButton.addEventListener('click', () => {
  runBuild();
});

function markStale() {
  if (!result) return;
  clearPreview();
  setStatus('Column choice changed. Click Build PivotTable workbook again.', true);
}

rowInput.addEventListener('change', markStale);
valueInput.addEventListener('change', markStale);

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
  const blob = new Blob([result.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = result.fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

refreshFields();
downloadButton.disabled = true;
