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

const server = await startServer();
const profile = await mkdtemp(join(tmpdir(), 'pivot-percent-change-'));
const xlsxPath = join(profile, 'periods.xlsx');
const badPath = join(profile, 'bad.xlsx');
const sheet = XLSX.utils.aoa_to_sheet([
  ['Region', 'Period', 'Amount'],
  ['East', '2024-01', 10],
  ['East', '2024-03', 40],
  ['West', '2024-03', 'n/a'],
]);
const book = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(book, sheet, 'Sales');
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
    url: `http://127.0.0.1:${server.address().port}/pivot-percent-change.html`,
  });
  await loaded;

  const before = await evaluate(session, `(() => {
    const h1 = document.querySelectorAll('h1').length;
    document.getElementById('run').disabled = false;
    document.getElementById('run').click();
    return {
      h1,
      status: document.getElementById('tool-status').textContent,
      error: document.getElementById('tool-status').classList.contains('error'),
      hidden: document.getElementById('preview').hidden,
      rows: document.querySelectorAll('#preview tbody tr').length,
    };
  })()`);
  if (before.h1 !== 1 || !before.error || !before.hidden || before.rows !== 0 || !before.status.includes('before calculating')) {
    throw new Error(`Empty state showed a result: ${JSON.stringify(before)}`);
  }

  await evaluate(session, `document.getElementById('example').click()`);
  const exampleRows = await waitFor(session, `(() => {
    if (document.querySelector('#preview')?.dataset.ready !== '1') return '';
    return JSON.stringify([...document.querySelectorAll('#preview tbody tr')].map((tr) => [...tr.cells].map((cell) => cell.textContent)));
  })()`);
  const example = JSON.parse(exampleRows);
  const expectedExample = [
    ['East', '2024-01', '50', '25', '', 'n/a', 'n/a'],
    ['East', '2024-02', '80', '40', '2024-01', '30', '60'],
    ['East', '2024-04', '20', '10', '2024-02', '-60', '-75'],
    ['West', '2024-01', '', 'n/a', '', 'n/a', 'n/a'],
    ['West', '2024-04', '50', '25', '2024-02', 'n/a', 'n/a'],
    ['North', '2024-02', '0', '0', '2024-01', 'n/a', 'n/a'],
  ];
  if (JSON.stringify(example) !== JSON.stringify(expectedExample)) {
    throw new Error(`Example table mismatch: ${exampleRows}`);
  }
  if (exampleRows.includes('NaN') || exampleRows.includes('Infinity') || exampleRows.includes('2024-03')) {
    throw new Error(`Example invented a number or a missing month: ${exampleRows}`);
  }

  const csv = await evaluate(session, `new Promise((resolve) => {
    const orig = URL.createObjectURL;
    URL.createObjectURL = function(blob) {
      blob.text().then(resolve);
      return orig.call(URL, blob);
    };
    document.getElementById('download').click();
  })`);
  if (!csv.includes('East,2024-02,80,40,2024-01,30,60') || !csv.includes('West,2024-04,50,25,2024-02,n/a,n/a')) {
    throw new Error(`Download CSV missing calculated rows: ${csv}`);
  }

  await evaluate(session, `document.getElementById('source').value = 'Region,Period,Amount\\nEast,2024-01,5\\nEast,2024-02,-5'; document.getElementById('file').value = ''; document.getElementById('load').click();`);
  const zeroRows = JSON.parse(await waitFor(session, `(() => {
    const status = document.getElementById('tool-status').textContent || '';
    if (!status.includes('Grand total is 0')) return '';
    return JSON.stringify([...document.querySelectorAll('#preview tbody tr')].map((tr) => [...tr.cells].map((cell) => cell.textContent)));
  })()`));
  if (JSON.stringify(zeroRows) !== JSON.stringify([
    ['East', '2024-01', '5', 'n/a', '', 'n/a', 'n/a'],
    ['East', '2024-02', '-5', 'n/a', '2024-01', '-10', '-200'],
  ]) || JSON.stringify(zeroRows).includes('NaN')) {
    throw new Error(`Zero denominator invented a share: ${JSON.stringify(zeroRows)}`);
  }

  await evaluate(session, `document.getElementById('source').value = 'Region,Period,Amount\\nEast,2024-01,\\nEast,2024-03,8'; document.getElementById('load').click();`);
  const missingRows = JSON.parse(await waitFor(session, `(() => {
    if (document.querySelector('#preview')?.dataset.ready !== '1') return '';
    const rows = [...document.querySelectorAll('#preview tbody tr')].map((tr) => [...tr.cells].map((cell) => cell.textContent));
    return rows[1] && rows[1][1] === '2024-03' ? JSON.stringify(rows) : '';
  })()`));
  if (JSON.stringify(missingRows) !== JSON.stringify([
    ['East', '2024-01', '', 'n/a', '', 'n/a', 'n/a'],
    ['East', '2024-03', '8', '100', '2024-01', 'n/a', 'n/a'],
  ])) {
    throw new Error(`Missing previous period was filled in: ${JSON.stringify(missingRows)}`);
  }

  await session.send('DOM.enable');
  const doc = await session.send('DOM.getDocument', { depth: 1, pierce: true });
  const input = await session.send('DOM.querySelector', { nodeId: doc.root.nodeId, selector: '#file' });
  await session.send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: [xlsxPath] });
  await evaluate(session, `document.getElementById('load').click()`);
  const xlsxRows = JSON.parse(await waitFor(session, `(() => {
    const status = document.getElementById('tool-status').textContent || '';
    if (!status.startsWith('Worksheet Sales:')) return '';
    return JSON.stringify([...document.querySelectorAll('#preview tbody tr')].map((tr) => [...tr.cells].map((cell) => cell.textContent)));
  })()`));
  if (JSON.stringify(xlsxRows) !== JSON.stringify([
    ['East', '2024-01', '10', '20', '', 'n/a', 'n/a'],
    ['East', '2024-03', '40', '80', '2024-01', '30', '300'],
    ['West', '2024-03', '', 'n/a', '2024-01', 'n/a', 'n/a'],
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

  console.log('pivot percent change UI: browser checks passed');
} finally {
  if (session) session.close();
  if (chrome) chrome.kill('SIGKILL');
  server.close();
  await rm(profile, { recursive: true, force: true });
}
