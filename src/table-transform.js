export function transposeTable(table) {
  if (table == null || table.length === 0) {
    throw new Error('Cannot transpose an empty table');
  }

  const numRows = table.length;
  let maxCols = 0;
  for (let r = 0; r < numRows; r++) {
    const row = table[r];
    if (row == null) {
      throw new Error('Invalid table row');
    }
    maxCols = Math.max(maxCols, row.length);
  }

  const out = [];
  for (let c = 0; c < maxCols; c++) {
    const newRow = [];
    for (let r = 0; r < numRows; r++) {
      const row = table[r];
      newRow.push(c < row.length ? row[c] : '');
    }
    out.push(newRow);
  }
  return out;
}

export function mergeTables(tables) {
  if (tables == null || tables.length < 2) {
    throw new Error('mergeTables requires at least two tables');
  }

  for (const table of tables) {
    if (table == null || table.length === 0) {
      throw new Error('Each table must be nonempty');
    }
    const width = table[0].length;
    if (width === 0) {
      throw new Error('Header row must not be empty');
    }
    for (let i = 0; i < table.length; i++) {
      if (table[i].length !== width) {
        throw new Error('Inconsistent row lengths within a table');
      }
    }
    const headerRow = table[0];
    const seen = new Set();
    for (const h of headerRow) {
      if (h === '') {
        throw new Error('Blank header cell');
      }
      if (seen.has(h)) {
        throw new Error('Duplicate header in table');
      }
      seen.add(h);
    }
  }

  const unionHeaders = [];
  const unionSeen = new Set();
  for (const table of tables) {
    for (const h of table[0]) {
      if (!unionSeen.has(h)) {
        unionSeen.add(h);
        unionHeaders.push(h);
      }
    }
  }

  const merged = [unionHeaders.slice()];

  for (const table of tables) {
    const headerRow = table[0];
    const indexByHeader = new Map();
    for (let c = 0; c < headerRow.length; c++) {
      indexByHeader.set(headerRow[c], c);
    }

    for (let r = 1; r < table.length; r++) {
      const srcRow = table[r];
      const outRow = unionHeaders.map((h) => {
        const idx = indexByHeader.get(h);
        return idx !== undefined ? srcRow[idx] : '';
      });
      merged.push(outRow);
    }
  }

  return merged;
}
