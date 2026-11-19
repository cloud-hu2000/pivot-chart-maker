import {
  EXAMPLE_CONDITIONS,
  EXAMPLE_CSV,
  LIMITS,
  OPERATORS,
  PREVIEW_ROW_LIMIT,
  filterRows,
  openXlsx,
  parseCsv,
  previewMatches,
  serializeCsv,
} from './csv-filter.js';

const fileInput = document.querySelector('#file');
const delimiterInput = document.querySelector('#delimiter');
const sheetWrap = document.querySelector('#sheet-wrap');
const sheetInput = document.querySelector('#sheet');
const sourceInput = document.querySelector('#source');
const modeInput = document.querySelector('#mode');
const conditionList = document.querySelector('#conditions');
const addButton = document.querySelector('#add-condition');
const loadButton = document.querySelector('#load');
const exampleButton = document.querySelector('#example');
const runButton = document.querySelector('#run');
const downloadButton = document.querySelector('#download');
const statusNode = document.querySelector('#tool-status');
const preview = document.querySelector('#preview');

let table = null;
let workbook = null;
let result = null;
let suppressStale = false;

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
  refreshColumns();
  setStatus(error instanceof Error ? error.message : 'The file could not be read.', true);
}

function showFilterError(error) {
  clearPreview();
  setStatus(error instanceof Error ? error.message : 'The rows could not be filtered.', true);
}

function addOption(select, value, label) {
  const option = document.createElement('option');
  option.value = value;
  option.textContent = label;
  select.appendChild(option);
}

function operatorSelect() {
  const select = document.createElement('select');
  select.className = 'cond-op';
  for (const operator of OPERATORS) addOption(select, operator.id, operator.label);
  return select;
}

function syncValueField(row) {
  const operator = OPERATORS.find((item) => item.id === row.querySelector('.cond-op').value);
  const input = row.querySelector('.cond-value');
  const unused = !operator || operator.kind === 'empty';
  input.disabled = unused;
  input.placeholder = unused ? 'Not used for this operator' : 'Comparison value';
  if (unused) input.value = '';
}

function refreshColumns() {
  const headers = table ? table.headers : [];
  for (const select of conditionList.querySelectorAll('.cond-column')) {
    const previous = select.value;
    select.replaceChildren();
    if (!headers.length) {
      addOption(select, '', 'Load data first');
      select.disabled = true;
      select.value = '';
      continue;
    }
    headers.forEach((name, index) => addOption(select, String(index), name));
    select.disabled = false;
    select.value = [...select.options].some((option) => option.value === previous) ? previous : '0';
  }
}

function addCondition(preset) {
  const row = document.createElement('div');
  row.className = 'condition';
  const columnLabel = document.createElement('label');
  columnLabel.textContent = 'Column';
  const column = document.createElement('select');
  column.className = 'cond-column';
  columnLabel.appendChild(column);
  const operatorLabel = document.createElement('label');
  operatorLabel.textContent = 'Operator';
  const operator = operatorSelect();
  operatorLabel.appendChild(operator);
  const valueLabel = document.createElement('label');
  valueLabel.textContent = 'Value';
  const value = document.createElement('input');
  value.className = 'cond-value';
  value.type = 'text';
  valueLabel.appendChild(value);
  const remove = document.createElement('button');
  remove.className = 'secondary cond-remove';
  remove.type = 'button';
  remove.textContent = 'Remove';
  row.append(columnLabel, operatorLabel, valueLabel, remove);
  conditionList.appendChild(row);
  refreshColumns();
  if (preset) {
    operator.value = preset.operator;
    if (table && Number.isInteger(Number(preset.column))) column.value = String(preset.column);
    value.value = preset.value == null ? '' : String(preset.value);
  }
  syncValueField(row);
  return row;
}

function readConditions() {
  return [...conditionList.querySelectorAll('.condition')].map((row) => ({
    column: row.querySelector('.cond-column').value,
    operator: row.querySelector('.cond-op').value,
    value: row.querySelector('.cond-value').value,
  }));
}

function renderResult(filtered) {
  result = filtered;
  const visible = previewMatches(filtered, PREVIEW_ROW_LIMIT);
  preview.replaceChildren();
  const tableNode = document.createElement('table');
  const caption = document.createElement('caption');
  caption.textContent = filtered.mode === 'and' ? 'Rows matching every condition' : 'Rows matching any condition';
  tableNode.appendChild(caption);
  const head = document.createElement('thead');
  const headRow = document.createElement('tr');
  for (const header of filtered.headers) {
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
  preview.dataset.matchCount = String(filtered.matchCount);
  downloadButton.disabled = false;
  const sheetName = table && table.name ? `Worksheet ${table.name}: ` : '';
  const dataNoun = filtered.dataRowCount === 1 ? 'row' : 'rows';
  let detail;
  if (filtered.matchCount === 0) {
    detail = 'The download contains the header and no data rows.';
  } else if (filtered.rows.length > visible.length) {
    detail = `The preview shows the first ${PREVIEW_ROW_LIMIT} matches in file order. The download contains all ${filtered.matchCount} matches.`;
  } else {
    detail = 'The preview lists every match. The download contains only those rows.';
  }
  setStatus(`${sheetName}${filtered.matchCount} of ${filtered.dataRowCount} data ${dataNoun} matched. ${detail}`, false);
}

function runFilter() {
  if (!table) {
    showFilterError(new Error('Choose a CSV or XLSX file, or paste CSV, before filtering.'));
    return;
  }
  try {
    renderResult(filterRows(table, readConditions(), modeInput.value));
  } catch (error) {
    showFilterError(error);
  }
}

function markStale() {
  if (suppressStale || !result) return;
  clearPreview();
  setStatus('Conditions changed. Click Filter rows to update the matches.', true);
}

async function useTable(nextTable) {
  table = nextTable;
  clearPreview();
  refreshColumns();
  const sheetName = table.name ? `Worksheet ${table.name}: ` : '';
  const noun = table.rows.length === 1 ? 'row' : 'rows';
  setStatus(`${sheetName}Loaded ${table.rows.length} data ${noun}. Add a condition, then click Filter rows.`, false);
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

function applyExampleConditions() {
  suppressStale = true;
  conditionList.replaceChildren();
  modeInput.value = 'and';
  for (const condition of EXAMPLE_CONDITIONS) addCondition(condition);
  suppressStale = false;
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
  applyExampleConditions();
  runFilter();
});

runButton.addEventListener('click', () => {
  runFilter();
});

addButton.addEventListener('click', () => {
  addCondition();
  markStale();
});

modeInput.addEventListener('change', () => {
  markStale();
});

conditionList.addEventListener('input', () => {
  markStale();
});

conditionList.addEventListener('change', (event) => {
  const row = event.target.closest('.condition');
  if (row && event.target.classList.contains('cond-op')) syncValueField(row);
  markStale();
});

conditionList.addEventListener('click', (event) => {
  const button = event.target.closest('.cond-remove');
  if (!button) return;
  button.closest('.condition').remove();
  markStale();
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
  link.download = 'csv-filter-matches.csv';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
});

addCondition();
runButton.disabled = false;
downloadButton.disabled = true;
