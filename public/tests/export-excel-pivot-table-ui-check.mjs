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
        close() { ws.close(); },
      });
    });
    ws.addEventListener('message', (event) => {
      const message = JSON.parse(event.data);
      if (message.id && pending.has(message.id)) {
        const { res, rej } = pending.get(message.id);
        pending.delete(message.id);
        if (message.error) rej(new Error(JSON.stringify(message.error)));
        else res(message.result || {});
      }
      if (message.method === 'Page.loadEventFired') {
        const waiter = loadWaiters.shift();
        if (waiter) waiter();
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

const server = await startServer();
const profile = await mkdtemp(join(tmpdir(), 'excel-pivot-'));
const xlsxPath = join(profile, 'rows.xlsx');
const badPath = join(profile, 'bad.xlsx');
const sheet = XLSX.utils.aoa_to_sheet([
  ['Region', 'Product', 'Amount'],
  ['East', 'Desk', 10],
  ['West', 'Desk', 4],
]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, sheet, 'Sales');
XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
  ['Region', 'Amount'],
  ['North', 3],
]), 'Other');
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
    url: `http://127.0.0.1:${server.address().port}/export-excel-pivot-table.html`,
  });
  await loaded;
  await waitFor(session, `document.querySelector('#run') ? 'ready' : ''`);

  const before = await evaluate(session, `(() => {
    const h1 = document.querySelectorAll('h1').length;
    document.getElementById('run').click();
    return {
      h1,
      status: document.getElementById('tool-status').textContent,
      error: document.getElementById('tool-status').classList.contains('error'),
      hidden: document.getElementById('preview').hidden,
      rows: document.querySelectorAll('#preview tbody tr').length,
    };
  })()`);
  if (before.h1 !== 1 || !before.error || !before.hidden || before.rows !== 0 || !before.status.includes('before building')) {
    throw new Error(`Empty state showed a result: ${JSON.stringify(before)}`);
  }

  await evaluate(session, `document.getElementById('example').click()`);
  const exampleRows = await waitFor(session, `(() => {
    if (document.querySelector('#preview')?.dataset.ready !== '1') return '';
    if (document.querySelector('#preview')?.dataset.groupCount !== '3') return '';
    return JSON.stringify([...document.querySelectorAll('#preview tbody tr')].map((tr) => [...tr.cells].map((cell) => cell.textContent)));
  })()`);
  const expected = JSON.stringify([
    ['East', '15', '2', '0'],
    ['West', '7', '1', '1'],
    ['west', '', '0', '1'],
    ['Grand Total', '22', '3', '2'],
  ]);
  if (exampleRows !== expected) throw new Error(`Example table mismatch: ${exampleRows}`);
  const exampleStatus = await evaluate(session, `document.getElementById('tool-status').textContent`);
  if (!exampleStatus.includes('3 groups') || !exampleStatus.includes('22') || exampleStatus.includes('NaN') || /refreshable|refreshes on open/i.test(exampleStatus)) {
    throw new Error(`Example status mismatch: ${exampleStatus}`);
  }

  const download = await evaluate(session, `new Promise((resolve) => {
    const orig = URL.createObjectURL;
    URL.createObjectURL = function(blob) {
      blob.arrayBuffer().then((buffer) => {
        const bytes = new Uint8Array(buffer);
        resolve(JSON.stringify({
          size: bytes.length,
          pk: bytes[0] === 0x50 && bytes[1] === 0x4b,
          type: blob.type,
        }));
      });
      return orig.call(URL, blob);
    };
    document.getElementById('download').click();
  })`);
  const file = JSON.parse(download);
  if (!file.pk || file.size < 1000 || !file.type.includes('spreadsheetml')) {
    throw new Error(`Download was not an XLSX package: ${download}`);
  }

  await evaluate(session, `(() => {
    document.getElementById('row-field').value = '1';
    document.getElementById('row-field').dispatchEvent(new Event('change'));
  })()`);
  const stale = await evaluate(session, `(() => ({
    error: document.getElementById('tool-status').classList.contains('error'),
    hidden: document.getElementById('preview').hidden,
    disabled: document.getElementById('download').disabled,
    status: document.getElementById('tool-status').textContent,
  }))()`);
  if (!stale.error || !stale.hidden || !stale.disabled || !stale.status.includes('Column choice changed')) {
    throw new Error(`Stale column change kept the old workbook: ${JSON.stringify(stale)}`);
  }

  await session.send('DOM.enable');
  const doc = await session.send('DOM.getDocument', { depth: 1, pierce: true });
  const input = await session.send('DOM.querySelector', { nodeId: doc.root.nodeId, selector: '#file' });
  await session.send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: [xlsxPath] });
  await waitFor(session, `document.getElementById('tool-status').textContent.startsWith('Worksheet Sales:') ? 'loaded' : ''`);
  await evaluate(session, `(() => {
    document.getElementById('row-field').value = '0';
    document.getElementById('value-field').value = '2';
    document.getElementById('run').click();
  })()`);
  const uploaded = await waitFor(session, `(() => {
    const status = document.getElementById('tool-status').textContent || '';
    if (!status.includes('Grand total of plain numbers is 14')) return '';
    return JSON.stringify([...document.querySelectorAll('#preview tbody tr')].map((tr) => [...tr.cells].map((cell) => cell.textContent)));
  })()`);
  if (uploaded !== JSON.stringify([['East', '10', '1', '0'], ['West', '4', '1', '0'], ['Grand Total', '14', '2', '0']])) {
    throw new Error(`Uploaded XLSX mismatch: ${uploaded}`);
  }

  await evaluate(session, `document.getElementById('sheet').value = '1'; document.getElementById('sheet').dispatchEvent(new Event('change'));`);
  await waitFor(session, `document.getElementById('tool-status').textContent.startsWith('Worksheet Other:') ? 'loaded' : ''`);
  await evaluate(session, `(() => {
    document.getElementById('row-field').value = '0';
    document.getElementById('value-field').value = '1';
    document.getElementById('run').click();
  })()`);
  const other = await waitFor(session, `(() => {
    const status = document.getElementById('tool-status').textContent || '';
    if (!status.includes('Grand total of plain numbers is 3')) return '';
    return status;
  })()`);
  if (!other.includes('1 group') || !other.includes('3')) {
    throw new Error(`Second sheet mismatch: ${other}`);
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

  console.log('excel pivot table UI: browser checks passed');
} finally {
  if (session) {
    try { session.close(); } catch { /* The page socket can already be closed. */ }
  }
  if (chrome) chrome.kill('SIGKILL');
  server.close();
  await new Promise((resolve) => setTimeout(resolve, 200));
  await rm(profile, { recursive: true, force: true }).catch(() => {});
}
