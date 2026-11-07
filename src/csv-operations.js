export function countValues(rows, column) {
  const counts = new Map();
  const order = [];
  const positions = new Map();

  for (const row of rows) {
    const value = row[column] ?? '';
    if (!counts.has(value)) {
      positions.set(value, order.length);
      order.push(value);
      counts.set(value, 0);
    }
    counts.set(value, counts.get(value) + 1);
  }

  const total = rows.length;
  const sorted = [...order].sort((a, b) => {
    const diff = counts.get(b) - counts.get(a);
    if (diff !== 0) return diff;
    return positions.get(a) - positions.get(b);
  });

  const outRows = sorted.map((value) => {
    const count = counts.get(value);
    const percent =
      total === 0 ? '0.00' : ((count / total) * 100).toFixed(2);
    return [value, String(count), percent];
  });

  return {
    headers: ['Value', 'Count', 'Percent'],
    rows: outRows,
  };
}

function uniqueHeaderName(base, used) {
  if (!used.has(base)) {
    used.add(base);
    return base;
  }
  let n = 2;
  while (used.has(`${base}_${n}`)) {
    n += 1;
  }
  const name = `${base}_${n}`;
  used.add(name);
  return name;
}

export function dedupe(rows, columns) {
  if (!columns || columns.length === 0) {
    throw new Error('columns must not be empty');
  }

  const seen = new Set();
  const outRows = [];
  let removed = 0;

  for (const row of rows) {
    const key = JSON.stringify(columns.map((i) => row[i] ?? ''));
    if (seen.has(key)) {
      removed += 1;
      continue;
    }
    seen.add(key);
    outRows.push(row.map((cell) => cell ?? ''));
  }

  return {
    rows: outRows,
    removed,
  };
}

export function unpivot(headers, rows, columns) {
  if (!columns || columns.length === 0) {
    throw new Error('columns must not be empty');
  }
  if (columns.length >= headers.length) {
    throw new Error('cannot unpivot all columns');
  }

  const colSet = new Set(columns);
  const keepIdx = [];
  for (let i = 0; i < headers.length; i += 1) {
    if (!colSet.has(i)) keepIdx.push(i);
  }

  const used = new Set(headers.map((h) => h ?? ''));
  const columnHeader = uniqueHeaderName('Column', used);
  const valueHeader = uniqueHeaderName('Value', used);

  const outHeaders = [
    ...keepIdx.map((i) => headers[i] ?? ''),
    columnHeader,
    valueHeader,
  ];

  const outRows = [];
  for (const row of rows) {
    const base = keepIdx.map((i) => row[i] ?? '');
    for (const ci of columns) {
      outRows.push([
        ...base,
        headers[ci] ?? '',
        row[ci] ?? '',
      ]);
    }
  }

  return {
    headers: outHeaders,
    rows: outRows,
  };
}