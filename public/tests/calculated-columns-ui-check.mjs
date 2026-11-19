import { spawn } from 'node:child_process';
import { createServer } from 'node:http';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from '@e965/xlsx';

const publicRoot = fileURLToPath(new URL('..', import.meta.url));
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
};

function startServer() {
  const server = createServer(async (req, res) => {
    try {
      const url = new URL(req.url, 'http://127.0.0.1');
      const rel = decodeURIComponent(url.pathname).replace(/^\/+/, '');
      const file = normalize(join(publicRoot, rel));
      if (file !== publicRoot && !file.startsWith(publicRoot)) {
        res.writeHead(403);
        res.end('forbidden');
        return;
      }
      const body = await readFile(file);
      res.writeHead(200, { 'content-type': types[extname(file)] || 'application/octet-stream', 'cache-control': 'no-store' });
      res.end(body);
    } catch {
      res.writeHead(404);
      res.end('not found');
    }
  });
  return new Promise((resolve) => {
    server.listen(0, '127.0.0.1', () => resolve(server));
  });
}

async function launchChrome(profile) {
  const chrome = spawn('google-chrome', [
    '--headless=new',
    '--no-sandbox',
    '--disable-gpu',
    '--disable-dev-shm-usage',
    '--remote-debugging-port=0',
    `--user-data-dir=${profile}`,
    'about:blank',
  ], { stdio: ['ignore', 'pipe', 'pipe'] });
  let output = '';
  chrome.stdout.on('data', (chunk) => { output += chunk; });
  chrome.stderr.on('data', (chunk) => { output += chunk; });
  const started = Date.now();
  while (Date.now() - started < 15000) {
    const match = output.match(/DevTools listening on ws:\/\/127\.0\.0\.1:(\d+)\//);
    if (match) return { chrome, port: match[1] };
    if (chrome.exitCode != null) break;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  chrome.kill('SIGKILL');
  throw new Error(`Chrome did not open a debugging port. Output: ${output.slice(0, 500)}`);
}

function connect(wsUrl) {
  const ws = new WebSocket(wsUrl);
  let seq = 0;
  const pending = new Map();
  const loadWaiters = [];
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('DevTools socket timed out')), 10000);
    ws.addEventListener('open', () => {
      clearTimeout(timer);
      resolve({
        send(method, params = {}) {
          const id = ++seq;
          ws.send(JSON.stringify({ id, method, params }));
          return new Promise((res, rej) => pending.set(id, { res, rej }));
        },
        nextLoad() {
          return new Promise((res) => loadWaiters.push(res));
        },
        close() {
          ws.close();
        },
      });
    });
    ws.addEventListener('message', (event) => {
      const msg = JSON.parse(event.data);
      if (msg.id && pending.has(msg.id)) {
        const { res, rej } = pending.get(msg.id);
        pending.delete(msg.id);
        if (msg.error) rej(new Error(JSON.stringify(msg.error)));
        else res(msg.result);
      }
      if (msg.method === 'Page.loadEventFired') {
        while (loadWaiters.length) loadWaiters.shift()();
      }
    });
    ws.addEventListener('error', () => {
      clearTimeout(timer);
      reject(new Error('DevTools socket failed'));
    });
  });
}

async function evaluate(session, expression) {
  const result = await session.send('Runtime.evaluate', {
    expression,
    awaitPromise: true,
    returnByValue: true,
  });
  if (result.exceptionDetails) {
    throw new Error(result.exceptionDetails.text || JSON.stringify(result.exceptionDetails));
  }
  return result.result?.value;
}

async function waitFor(session, expression) {
  const started = Date.now();
  let last = '';
  while (Date.now() - started < 8000) {
    last = await evaluate(session, expression);
    if (last) return last;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error(`Timed out waiting for page state. Last value: ${last}`);
}

function rowsExpression(readyCheck) {
  return `(() => {
    const preview = document.querySelector('#preview');
    if (preview?.dataset.ready !== '1') return '';
    if (!(${readyCheck})) return '';
    return JSON.stringify([...document.querySelectorAll('#preview tbody tr')].map((tr) => [...tr.cells].map((cell) => cell.textContent)));
  })()`;
}

const server = await startServer();
const profile = await mkdtemp(join(tmpdir(), 'calculated-columns-'));
const xlsxPath = join(profile, 'rows.xlsx');
const badPath = join(profile, 'bad.xlsx');
const sheet = XLSX.utils.aoa_to_sheet([
  ['Item', 'Width', 'Height'],
  ['tile', 2, 3],
  ['label', 'n/a', 1],
]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, sheet, 'Boxes');
await writeFile(xlsxPath, XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
await writeFile(badPath, 'this is not a workbook');

let chrome;
let session;
try {
  const launched = await launchChrome(profile);
  chrome = launched.chrome;
  const targets = await (await fetch(`http://127.0.0.1:${launched.port}/json/list`)).json();
  const page = targets.find((target) => target.type === 'page');
  if (!page) throw new Error('Chrome did not open a page target');
  session = await connect(page.webSocketDebuggerUrl);
  await session.send('Page.enable');
  const loaded = session.nextLoad();
  await session.send('Page.navigate', {
    url: `http://127.0.0.1:${server.address().port}/calculated-columns.html`,
  });
  await loaded;
  await waitFor(session, `document.querySelector('#column-names li') ? 'ready' : ''`);

  const before = await evaluate(session, `(() => {
    const h1 = document.querySelectorAll('h1').length;
    document.getElementById('run').click();
    return {
      h1,
      status: document.getElementById('tool-status').textContent,
      error: document.getElementById('tool-status').classList.contains('error'),
      hidden: document.getElementById('preview').hidden,
      rows: document.querySelectorAll('#preview tbody tr').length,
      names: document.querySelector('#column-names').textContent,
    };
  })()`);
  if (before.h1 !== 1 || !before.error || !before.hidden || before.rows !== 0 || !before.status.includes('before adding a column') || !before.names.includes('Load data')) {
    throw new Error(`Empty state showed a result: ${JSON.stringify(before)}`);
  }

  await evaluate(session, `document.getElementById('example').click()`);
  const exampleRows = JSON.parse(await waitFor(session, rowsExpression(`preview.dataset.calculatedCount === '3' && preview.dataset.rowCount === '5'`)));
  if (JSON.stringify(exampleRows) !== JSON.stringify([
    ['tile', '2', '3', '8'],
    ['panel', '10', '4', '18'],
    ['label', 'n/a', '1', ''],
    ['spacer', '', '5', ''],
    ['strip', '1.5', '2', '5.5'],
  ])) {
    throw new Error(`Example table mismatch: ${JSON.stringify(exampleRows)}`);
  }
  const exampleStatus = await evaluate(session, `document.getElementById('tool-status').textContent`);
  if (!exampleStatus.includes('3 of 5') || !exampleStatus.includes('not a plain number') || exampleStatus.includes('NaN') || exampleStatus.includes('Infinity')) {
    throw new Error(`Example status mismatch: ${exampleStatus}`);
  }

  const csv = await evaluate(session, `new Promise((resolve) => {
    const orig = URL.createObjectURL;
    URL.createObjectURL = function(blob) {
      blob.text().then(resolve);
      return orig.call(URL, blob);
    };
    document.getElementById('download').click();
  })`);
  if (!csv.includes('tile,2,3,8') || !csv.includes('label,n/a,1,') || !csv.includes('strip,1.5,2,5.5') || csv.includes('Infinity') || csv.includes('NaN') || csv.includes('=')) {
    throw new Error(`Download was not static row values: ${csv}`);
  }

  const stale = await evaluate(session, `(() => {
    document.getElementById('formula').value = '[Width] * [Height]';
    document.getElementById('formula').dispatchEvent(new Event('input'));
    return {
      hidden: document.getElementById('preview').hidden,
      disabled: document.getElementById('download').disabled,
      rows: document.querySelectorAll('#preview tbody tr').length,
    };
  })()`);
  if (!stale.hidden || !stale.disabled || stale.rows !== 0) {
    throw new Error(`Formula edit kept a stale download: ${JSON.stringify(stale)}`);
  }

  await evaluate(session, `document.getElementById('run').click()`);
  const productRows = JSON.parse(await waitFor(session, rowsExpression(`preview.dataset.calculatedCount === '3'`)));
  if (JSON.stringify(productRows) !== JSON.stringify([
    ['tile', '2', '3', '6'],
    ['panel', '10', '4', '40'],
    ['label', 'n/a', '1', ''],
    ['spacer', '', '5', ''],
    ['strip', '1.5', '2', '3'],
  ])) {
    throw new Error(`Product table mismatch: ${JSON.stringify(productRows)}`);
  }

  const replaced = await evaluate(session, `(() => {
    document.getElementById('column-name').value = 'Width';
    document.getElementById('run').click();
    return {
      status: document.getElementById('tool-status').textContent,
      error: document.getElementById('tool-status').classList.contains('error'),
      hidden: document.getElementById('preview').hidden,
      rows: document.querySelectorAll('#preview tbody tr').length,
    };
  })()`);
  if (!replaced.error || !replaced.hidden || replaced.rows !== 0 || !replaced.status.includes('does not replace')) {
    throw new Error(`Existing column name was not rejected: ${JSON.stringify(replaced)}`);
  }

  const badFormula = await evaluate(session, `(() => {
    document.getElementById('column-name').value = 'Calculated';
    document.getElementById('formula').value = 'SUM([Width])';
    document.getElementById('run').click();
    return {
      status: document.getElementById('tool-status').textContent,
      error: document.getElementById('tool-status').classList.contains('error'),
      hidden: document.getElementById('preview').hidden,
    };
  })()`);
  if (!badFormula.error || !badFormula.hidden || !badFormula.status.includes('unexpected')) {
    throw new Error(`Function formula was accepted: ${JSON.stringify(badFormula)}`);
  }

  await evaluate(session, `(() => {
    document.getElementById('file').value = '';
    document.getElementById('delimiter').value = ',';
    document.getElementById('source').value = 'Item,Width,Height\\na,4,0\\n';
    document.getElementById('formula').value = '[Width] / [Height]';
    document.getElementById('column-name').value = 'Calculated';
    document.getElementById('load').click();
  })()`);
  await waitFor(session, `document.getElementById('tool-status').textContent.includes('Loaded 1 data row') ? 'loaded' : ''`);
  await evaluate(session, `document.getElementById('run').click()`);
  const divided = JSON.parse(await waitFor(session, rowsExpression(`preview.dataset.rowCount === '1' && preview.dataset.calculatedCount === '0'`)));
  if (JSON.stringify(divided) !== JSON.stringify([['a', '4', '0', '']])) {
    throw new Error(`Division row mismatch: ${JSON.stringify(divided)}`);
  }
  const divideStatus = await evaluate(session, `document.getElementById('tool-status').textContent`);
  if (!divideStatus.includes('divisor is zero') || divideStatus.includes('Infinity') || divideStatus.includes('NaN')) {
    throw new Error(`Division status mismatch: ${divideStatus}`);
  }

  await session.send('DOM.enable');
  const doc = await session.send('DOM.getDocument', { depth: 1, pierce: true });
  const input = await session.send('DOM.querySelector', { nodeId: doc.root.nodeId, selector: '#file' });
  await session.send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: [xlsxPath] });
  await evaluate(session, `(() => {
    document.getElementById('formula').value = '[Width] + [Height] * 2';
    document.getElementById('column-name').value = 'Calculated';
    document.getElementById('load').click();
  })()`);
  await waitFor(session, `document.getElementById('tool-status').textContent.startsWith('Worksheet Boxes:') ? 'loaded' : ''`);
  await evaluate(session, `document.getElementById('run').click()`);
  const xlsxRows = JSON.parse(await waitFor(session, rowsExpression(`document.getElementById('tool-status').textContent.includes('1 of 2')`)));
  if (JSON.stringify(xlsxRows) !== JSON.stringify([
    ['tile', '2', '3', '8'],
    ['label', 'n/a', '1', ''],
  ])) {
    throw new Error(`XLSX UI mismatch: ${JSON.stringify(xlsxRows)}`);
  }

  const docAgain = await session.send('DOM.getDocument', { depth: 1, pierce: true });
  const inputAgain = await session.send('DOM.querySelector', { nodeId: docAgain.root.nodeId, selector: '#file' });
  await session.send('DOM.setFileInputFiles', { nodeId: inputAgain.nodeId, files: [badPath] });
  await evaluate(session, `document.getElementById('load').click()`);
  const bad = await waitFor(session, `(() => {
    const status = document.getElementById('tool-status');
    if (!status.classList.contains('error')) return '';
    return JSON.stringify({ status: status.textContent, hidden: document.getElementById('preview').hidden, rows: document.querySelectorAll('#preview tbody tr').length });
  })()`);
  const badState = JSON.parse(bad);
  if (!badState.hidden || badState.rows !== 0 || !/could not be read/i.test(badState.status)) {
    throw new Error(`Bad XLSX did not fail cleanly: ${bad}`);
  }

  console.log('calculated columns UI: browser checks passed');
} finally {
  if (session) session.close();
  if (chrome) chrome.kill('SIGKILL');
  server.close();
  await rm(profile, { recursive: true, force: true });
}
