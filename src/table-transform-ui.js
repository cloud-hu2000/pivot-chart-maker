import Papa from 'papaparse';
import {transposeTable, mergeTables} from './table-transform.js';
import './tool-pages.css';
const $=s=>document.querySelector(s), mode=document.body.dataset.tableTool, file=$('#file'), source=$('#source'), sheets=$('#sheet'), status=$('#tool-status'), preview=$('#preview'), download=$('#download');
let output=null, workbook=null, base='result';
function say(s,error=false){status.textContent=s;status.classList.toggle('error',error)}
function clear(){output=null;download.disabled=true;preview.hidden=true;preview.textContent=''}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function parse(text){const p=Papa.parse(text.replace(/^\uFEFF/,''),{skipEmptyLines:'greedy',delimiter:$('#delimiter').value||undefined});const error=p.errors.find(e=>e.code!=='UndetectableDelimiter');if(error)throw new Error(`CSV parsing error: ${error.message}`);if(!p.data.length)throw new Error('The CSV has no rows.');if(p.data.length>100001||p.data.reduce((n,r)=>Math.max(n,r.length),0)>1000)throw new Error('CSV limit: 100,000 data rows and 1,000 columns.');return p.data;}
function display(table){if(!table.length||!table[0].length)throw new Error('The result is empty.');output=table;preview.innerHTML=`<table><tbody>${table.slice(0,12).map((r,i)=>`<tr>${r.map(c=>i===0?`<th>${esc(c)}</th>`:`<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;preview.hidden=false;download.disabled=false;say(`${table.length.toLocaleString()} output rows; preview shows up to 12. Download contains the full result.`)}
async function readFile(f){if(f.size>10*1024*1024)throw new Error(`${f.name} exceeds 10 MB.`);return f.text()}
file.addEventListener('change',()=>{clear();workbook=null;if(sheets){sheets.innerHTML='';sheets.disabled=true;}if(source)source.value='';say('File selected. Click the button below to process it.')});
if(source)source.addEventListener('input',()=>{file.value='';clear();say('Input changed. Run the tool again.')});
if($('#delimiter'))$('#delimiter').addEventListener('change',()=>{clear();say('Delimiter changed. Run the tool again.')});
if(sheets)sheets.addEventListener('change',()=>{clear();say('Worksheet changed. Convert again.')});
$('#run').addEventListener('click',async()=>{clear();try{
  if(mode==='transpose'){
   let f=file.files[0],text=f?await readFile(f):source.value;if(!text.trim())throw new Error('Choose a CSV or paste data.');if(new TextEncoder().encode(text).length>10*1024*1024)throw new Error('Input exceeds 10 MB.');let table=parse(text);if(table.length>2000)throw new Error('Transpose output exceeds 2,000 columns. Use 2,000 rows or fewer.');base=f?f.name.replace(/\.[^.]+$/,''):'transposed';display(transposeTable(table));
  } else if(mode==='merge'){
   const files=[...file.files];if(files.length<2)throw new Error('Choose at least two CSV files.');if(files.length>20)throw new Error('Choose at most 20 files.');const tables=[];for(const f of files)tables.push(parse(await readFile(f)));const rows=tables.reduce((n,t)=>n+t.length-1,0);const cols=new Set(tables.flatMap(t=>t[0])).size;if(rows>100000||cols>1000||rows*cols>5000000)throw new Error('Merged result exceeds 100,000 data rows, 1,000 columns, or 5 million cells.');base='merged';display(mergeTables(tables));
  } else {
   const f=file.files[0];if(!f)throw new Error('Choose an Excel workbook first.');if(f.size>10*1024*1024)throw new Error('Workbook exceeds 10 MB.');if(!/\.(xlsx|xls)$/i.test(f.name))throw new Error('Choose an .xlsx or .xls workbook.');
   if(!workbook){const XLSX=await import('@e965/xlsx');workbook=XLSX.read(await f.arrayBuffer(),{type:'array'});if(!workbook.SheetNames.length)throw new Error('This workbook has no worksheets.');sheets.innerHTML=workbook.SheetNames.map((n,i)=>`<option value="${i}">${esc(n)}</option>`).join('');sheets.disabled=false;}
   const XLSX=await import('@e965/xlsx'),sheet=workbook.Sheets[workbook.SheetNames[Number(sheets.value)]];if(!sheet['!ref'])throw new Error('The selected worksheet is empty.');const range=XLSX.utils.decode_range(sheet['!ref']);if(range.e.r-range.s.r+1>100001||range.e.c-range.s.c+1>1000)throw new Error('Worksheet exceeds 100,001 rows or 1,000 columns.');const table=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,defval:'',blankrows:true});base=f.name.replace(/\.[^.]+$/,'');display(table);
  }
 }catch(e){clear();say(e.message||'Could not process this file.',true)}});
download.addEventListener('click',()=>{if(!output)return;const url=URL.createObjectURL(new Blob(['\uFEFF'+Papa.unparse(output)],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=`${base}-${mode}.csv`;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000)});
