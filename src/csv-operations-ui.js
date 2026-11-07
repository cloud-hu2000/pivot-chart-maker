import Papa from 'papaparse';
import { countValues, dedupe, unpivot } from './csv-operations.js';
import './tool-pages.css';

const $ = (s) => document.querySelector(s);
const mode = document.body.dataset.csvOperation;
const file = $('#file'), input = $('#source'), fields = $('#columns'), run = $('#run'), download = $('#download'), status = $('#tool-status'), preview = $('#preview');
let headers = [], rows = [], result = null, basename = 'result';
const escape = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function message(s, error = false) { status.textContent = s; status.classList.toggle('error', error); }
function clear() { result = null; download.disabled = true; preview.hidden = true; preview.innerHTML = ''; }
function reset() { headers = []; rows = []; fields.innerHTML = ''; run.disabled = true; clear(); }
function show(data) {
  preview.innerHTML = `<table><thead><tr>${data.headers.map(x => `<th>${escape(x)}</th>`).join('')}</tr></thead><tbody>${data.rows.slice(0,12).map(row => `<tr>${row.map(x => `<td>${escape(x)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  preview.hidden = false;
}
function parse(text) {
  const parsed = Papa.parse(text.replace(/^\uFEFF/, ''), {skipEmptyLines:'greedy', delimiter: $('#delimiter').value || undefined});
  if (parsed.errors.length) throw new Error(`CSV error: ${parsed.errors[0].message}`);
  if (parsed.data.length < 2) throw new Error('Add a header and at least one data row.');
  const width = parsed.data[0].length;
  if (parsed.data.some(row => row.length !== width)) throw new Error('Rows have different numbers of fields. Check the delimiter and quotes.');
  if (width > 1000 || parsed.data.length > 100001) throw new Error('Limit: 100,000 data rows and 1,000 columns.');
  const used = new Set();
  headers = parsed.data[0].map((v, i) => {let base = v || `Column ${i+1}`, name = base, n = 2; while (used.has(name.toLowerCase())) name = `${base} ${n++}`; used.add(name.toLowerCase()); return name;});
  rows = parsed.data.slice(1);
  fields.innerHTML = headers.map((h, i) => `<label class="field-choice"><input type="${mode === 'count' ? 'radio' : 'checkbox'}" name="column" value="${i}" ${mode === 'count' && i === 0 ? 'checked' : ''}> ${escape(h)}</label>`).join('');
  run.disabled = false;
  message(`${rows.length.toLocaleString()} rows and ${headers.length} columns loaded. Choose columns and run the tool.`);
}
async function load() {
  reset();
  try { const f=file.files[0]; if (f && f.size > 10 * 1024 * 1024) throw new Error('Choose a file no larger than 10 MB.');
    let text = f ? await f.text() : input.value; if (!text.trim()) throw new Error('Choose a CSV file or paste CSV data.');
    if (new TextEncoder().encode(text).length > 10 * 1024 * 1024) throw new Error('Input is larger than 10 MB.');
    basename = f ? f.name.replace(/\.[^.]+$/, '') : 'result'; parse(text);
  } catch(e) {reset();message(e.message,true);}
}
$('#load').addEventListener('click', load);
file.addEventListener('change', () => { input.value = ''; reset(); message('File selected. Click Load data.'); });
input.addEventListener('input', () => { if (file.value) file.value=''; reset(); message('Data changed. Click Load data.'); });
$('#delimiter').addEventListener('change', () => { reset(); message('Delimiter changed. Click Load data.'); });
fields.addEventListener('change', () => { clear(); message('Selection changed. Run the tool to update the result.'); });
run.addEventListener('click', () => { clear(); try {
  const columns=[...fields.querySelectorAll('input:checked')].map(e=>Number(e.value));
  if(!columns.length) throw new Error('Select at least one column.');
  if(mode==='count') result=countValues(rows,columns[0]);
  if(mode==='dedupe') result={headers,...dedupe(rows,columns)};
  if(mode==='unpivot') { if(columns.length === headers.length) throw new Error('Keep at least one identifier column unselected.'); if(rows.length * columns.length > 300000) throw new Error('Output exceeds 300,000 rows. Choose fewer columns or a smaller file.'); result=unpivot(headers,rows,columns); }
  show(result);download.disabled=false;
  message(`${result.rows.length.toLocaleString()} output rows${mode==='dedupe' ? `; ${result.removed.toLocaleString()} duplicate rows removed` : ''}. Preview shows up to 12 rows; download includes all rows.`);
} catch(e) {message(e.message,true);} });
download.addEventListener('click', () => { if(!result)return; const url=URL.createObjectURL(new Blob(['\uFEFF'+Papa.unparse([result.headers,...result.rows])],{type:'text/csv;charset=utf-8'})); const a=document.createElement('a');a.href=url;a.download=`${basename}-${mode}.csv`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000); });
