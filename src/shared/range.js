// Utilities for handling Excel custom-functions inputs (scalar or 2D ranges).
//
// IMPORTANT: No hard row/column/character limits are applied by default.
// The only practical limits are the model context window / maxOutputTokens and Excel's own constraints.

export function isMatrix(v) {
  return Array.isArray(v) && (v.length === 0 || Array.isArray(v[0]));
}

export function cellToString(cell, { maxCellChars } = {}) {
  let s = "";
  if (cell === null || cell === undefined) s = "";
  else if (typeof cell === "string") s = cell;
  else if (typeof cell === "number" || typeof cell === "boolean") s = String(cell);
  else s = JSON.stringify(cell);

  // Optional per-cell clamp (useful to avoid accidentally pushing huge blobs into prompts).
  const limit = Number.isFinite(Number(maxCellChars)) && Number(maxCellChars) > 0 ? Math.floor(Number(maxCellChars)) : Infinity;
  if (s.length > limit) s = s.slice(0, limit);
  return s;
}

export function matrixToTSV(matrix, opts = {}) {
  if (!Array.isArray(matrix)) return "";

  const {
    delimiter = "\t",
    rowDelimiter = "\n",
    maxChars,
    maxRows,
    maxCols,
    maxCellChars
  } = opts;

  const rowLimit = Number.isFinite(Number(maxRows)) && Number(maxRows) > 0 ? Math.floor(Number(maxRows)) : Infinity;
  const colLimit = Number.isFinite(Number(maxCols)) && Number(maxCols) > 0 ? Math.floor(Number(maxCols)) : Infinity;
  const charLimit = Number.isFinite(Number(maxChars)) && Number(maxChars) > 0 ? Math.floor(Number(maxChars)) : Infinity;

  const lines = [];
  let totalLen = 0;
  const safeDelimiter = String(delimiter);
  const safeRowDelimiter = String(rowDelimiter);

  for (let r = 0; r < matrix.length && r < rowLimit; r++) {
    const row = Array.isArray(matrix[r]) ? matrix[r] : [matrix[r]];
    const cols = colLimit === Infinity ? row : row.slice(0, colLimit);
    const line = cols
      .map((c) => cellToString(c, { maxCellChars }).replace(/\t/g, " ").replace(/\n/g, " "))
      .join(safeDelimiter);

    const extra = (lines.length ? safeRowDelimiter.length : 0) + line.length;
    if (totalLen + extra > charLimit) break;

    lines.push(line);
    totalLen += extra;
  }

  return lines.join(safeRowDelimiter);
}

export function normalizeTextInput(v, opts = {}) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (isMatrix(v)) return matrixToTSV(v, opts);
  if (Array.isArray(v)) return v.map((x) => cellToString(x, opts)).join("\n");
  return cellToString(v, opts);
}

export function flattenToStringList(matrix) {
  if (!Array.isArray(matrix)) return [];
  if (!isMatrix(matrix)) return matrix.map((x) => cellToString(x));

  const out = [];
  for (const row of matrix) {
    for (const cell of row) out.push(cellToString(cell));
  }
  return out;
}
