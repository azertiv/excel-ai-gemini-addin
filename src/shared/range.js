import { LIMITS } from "./constants";

export function isMatrix(v) {
  return Array.isArray(v) && (v.length === 0 || Array.isArray(v[0]));
}

export function cellToString(cell) {
  if (cell === null || cell === undefined) return "";
  if (typeof cell === "string") return cell;
  if (typeof cell === "number" || typeof cell === "boolean") return String(cell);
  try {
    if (typeof cell === "object") {
      if (cell.error) return String(cell.error);
      if (cell.basicValue) return String(cell.basicValue);
      return JSON.stringify(cell);
    }
  } catch { /* ignore */ }
  return String(cell);
}

export function matrixToTSV(matrix, maxChars = LIMITS.MAX_CONTEXT_CHARS) {
  if (!isMatrix(matrix)) return "";
  const maxRows = 200;
  const maxCols = 30;
  let out = "";
  let truncated = false;

  for (let r = 0; r < matrix.length && r < maxRows; r++) {
    const row = Array.isArray(matrix[r]) ? matrix[r] : [matrix[r]];
    const cells = [];
    for (let c = 0; c < row.length && c < maxCols; c++) {
      let s = cellToString(row[c]);
      s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      s = s.replace(/\t/g, " ").replace(/\n/g, " ");
      if (s.length > 200) s = s.slice(0, 200) + "…";
      cells.push(s);
    }
    const line = cells.join("\t");
    if (out.length + line.length + 1 > maxChars) { truncated = true; break; }
    out += (out ? "\n" : "") + line;
  }

  if (truncated) {
    out = out.slice(0, Math.max(0, maxChars - 20));
    out += "\n…(truncated)";
  }
  return out;
}

export function normalizeTextInput(v, maxChars = LIMITS.MAX_INPUT_CHARS) {
  if (v === null || v === undefined) return "";
  let s = typeof v === "string" ? v : String(v);
  s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (s.length > maxChars) s = s.slice(0, maxChars) + "\n…(truncated)";
  return s;
}

export function flattenToStringList(v) {
  if (isMatrix(v)) {
    const arr = [];
    for (const row of v) {
      const r = Array.isArray(row) ? row : [row];
      for (const cell of r) {
        const s = cellToString(cell).trim();
        if (s) arr.push(s);
      }
    }
    return arr;
  }
  const s = cellToString(v).trim();
  return s ? [s] : [];
}
