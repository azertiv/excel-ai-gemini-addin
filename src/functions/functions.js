// src/functions/functions.js
/* global CustomFunctions */

import { geminiGenerate, geminiMinimalTest } from "../shared/gemini.js";
import { getApiKey } from "../shared/storage.js";
import { ERR, DEFAULTS, LIMITS } from "../shared/constants.js";

// ---------- helpers ----------

function safeString(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  try {
    return JSON.stringify(v);
  } catch {
    return String(v);
  }
}

function parseOptions(optionsJson) {
  if (!optionsJson) return {};
  if (typeof optionsJson === "object" && optionsJson !== null) return optionsJson; // allow object in tests
  if (typeof optionsJson !== "string") return {};
  const s = optionsJson.trim();
  if (!s) return {};
  try {
    const obj = JSON.parse(s);
    return obj && typeof obj === "object" ? obj : {};
  } catch {
    return {};
  }
}

function clamp(n, min, max, fallback) {
  const x = Number(n);
  if (!Number.isFinite(x)) return fallback;
  return Math.min(max, Math.max(min, x));
}

function errorCode(code) {
  return code || ERR.API_ERROR;
}

function normalizeNewlines(s) {
  return String(s || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

function matrixToTSV(matrix, maxChars = LIMITS.MAX_CONTEXT_CHARS) {
  if (!Array.isArray(matrix)) return "";
  let out = "";
  for (const row of matrix) {
    if (out.length >= maxChars) break;
    const r = Array.isArray(row) ? row : [row];
    const line = r.map((c) => safeString(c).replace(/\t/g, " ").replace(/\n/g, " ")).join("\t");
    if (out) out += "\n";
    out += line;
  }
  if (out.length > maxChars) out = out.slice(0, maxChars) + "\n…(truncated)";
  return out;
}

function flattenLabels(labelsOrRange) {
  if (labelsOrRange === null || labelsOrRange === undefined) return [];
  if (typeof labelsOrRange === "string") {
    // accept "A|B|C" or "A,B,C"
    const s = labelsOrRange.trim();
    if (!s) return [];
    const parts = s.includes("|") ? s.split("|") : s.split(/[,;\n]+/);
    return parts.map((x) => x.trim()).filter(Boolean);
  }
  if (Array.isArray(labelsOrRange)) {
    // matrix -> flatten
    const out = [];
    for (const row of labelsOrRange) {
      if (Array.isArray(row)) {
        for (const cell of row) {
          const v = safeString(cell).trim();
          if (v) out.push(v);
        }
      } else {
        const v = safeString(row).trim();
        if (v) out.push(v);
      }
    }
    return out;
  }
  const s = safeString(labelsOrRange).trim();
  return s ? [s] : [];
}

function safeJsonParse(s) {
  try {
    return { ok: true, value: JSON.parse(s) };
  } catch (e) {
    return { ok: false, error: e };
  }
}

function coerceToTextOrJoin2D(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (Array.isArray(v)) {
    // if matrix, join as TSV
    if (v.length === 0) return "";
    if (Array.isArray(v[0])) return matrixToTSV(v, LIMITS.MAX_INPUT_CHARS);
    return v.map((x) => safeString(x)).join("\n");
  }
  return safeString(v);
}

function truncateForCell(s, maxChars = LIMITS.MAX_CELL_CHARS) {
  const t = normalizeNewlines(s);
  if (t.length <= maxChars) return t;
  return t.slice(0, maxChars) + "\n…(truncated)";
}

// ---------- per-function prompting ----------

function sysAsk(lang = "fr") {
  return [
    "You are an assistant embedded in Microsoft Excel custom functions.",
    `Respond in ${lang}.`,
    "Return a concise answer suitable for a single Excel cell (1 to 10 short lines).",
    "No Markdown. No code fences. No surrounding quotes.",
    "If the question cannot be answered from the provided information, say so briefly and suggest what to add."
  ].join("\n");
}

function sysTranslate(targetLang) {
  return [
    "You are a translation engine.",
    `Translate the user text into ${targetLang}.`,
    "Return only the translated text. No quotes. No explanations."
  ].join("\n");
}

function sysClassify(labels, lang = "en") {
  return [
    "You are a strict classifier.",
    `Return exactly one label from this set: ${labels.join(" | ")}`,
    "If uncertain, return exactly: UNKNOWN",
    `Respond in ${lang}.`,
    "Return only the label."
  ].join("\n");
}

function sysClean(lang = "fr") {
  return [
    "You are a text normalizer for spreadsheet cells.",
    `Respond in ${lang}.`,
    "Return only the cleaned text.",
    "No quotes. No explanations."
  ].join("\n");
}

function sysSummarize(lang = "fr") {
  return [
    "You summarize text for a spreadsheet cell.",
    `Respond in ${lang}.`,
    "Return 3 to 7 bullet points.",
    "Use '-' as bullet prefix.",
    "No Markdown headers. No code fences."
  ].join("\n");
}

function sysExtract(fields, lang = "fr") {
  return [
    "You extract structured fields from unstructured text.",
    `Respond in ${lang}.`,
    "Return STRICT JSON only (no Markdown, no code fences).",
    "Return an object with ONLY these keys:",
    fields.join(", "),
    "If a value is missing, use null.",
    "Dates should be ISO 8601 if possible.",
    "Numbers must be numbers (no currency symbols)."
  ].join("\n");
}

function sysTable(lang = "fr", maxRows = LIMITS.MAX_TABLE_ROWS, maxCols = LIMITS.MAX_TABLE_COLS, headers = null) {
  const headerLine = Array.isArray(headers) && headers.length ? `Use these headers exactly: ${headers.join(", ")}` : "";
  return [
    "You generate a table for Excel.",
    `Respond in ${lang}.`,
    "Return STRICT JSON only (no Markdown, no code fences).",
    "Return an object with shape: {\"headers\": string[], \"rows\": any[][]}.",
    `The number of columns must be <= ${maxCols}.`,
    `The number of rows must be <= ${maxRows}.`,
    "The first row of the spilled output will be headers.",
    headerLine,
    "Cells must be scalars (string/number/boolean/null)."
  ].filter(Boolean).join("\n");
}

function sysFill(lang = "fr", maxRows = LIMITS.MAX_FILL_ROWS) {
  return [
    "You are filling spreadsheet cells based on examples.",
    `Respond in ${lang}.`,
    "Return STRICT JSON only (no Markdown, no code fences).",
    "Return an object with shape: {\"values\": string[]}.",
    `Return at most ${maxRows} values.`,
    "Return only the values, in order, one per target row.",
    "If a value is unknown, return an empty string for that row."
  ].join("\n");
}

// ---------- core call wrapper ----------

async function callGemini({ system, user, options }) {
  const opt = options || {};
  const temperature = typeof opt.temperature === "number" ? clamp(opt.temperature, 0, 1, DEFAULTS.temperature) : DEFAULTS.temperature;
  const maxTokens = typeof opt.maxTokens === "number" ? clamp(opt.maxTokens, 1, 2048, DEFAULTS.maxTokens) : DEFAULTS.maxTokens;

  const timeoutMs = typeof opt.timeoutMs === "number" ? clamp(opt.timeoutMs, 1000, 60000, DEFAULTS.timeoutMs) : DEFAULTS.timeoutMs;
  const retry = typeof opt.retry === "number" ? clamp(opt.retry, 0, 2, DEFAULTS.retry) : DEFAULTS.retry;

  const cacheMode = typeof opt.cache === "string" ? opt.cache : DEFAULTS.cache;
  const cacheTtlSec = typeof opt.cacheTtlSec === "number" ? clamp(opt.cacheTtlSec, 0, 24 * 3600, DEFAULTS.cacheTtlSec) : DEFAULTS.cacheTtlSec;

  const res = await geminiGenerate({
    model: opt.model,
    system,
    user,
    generationConfig: { temperature, maxOutputTokens: maxTokens },
    cache: cacheMode,
    cacheTtlSec,
    timeoutMs,
    retry
    // responseMimeType/responseJsonSchema supported in your gemini.js already if needed
  });

  return res;
}

// ---------- Custom Functions ----------

/**
 * =AI.KEY_STATUS()
 * Returns "OK" if the API key is present, else "MISSING".
 */
export async function KEY_STATUS() {
  try {
    const key = await getApiKey();
    return key ? "OK" : "MISSING";
  } catch {
    return "MISSING";
  }
}

/**
 * =AI.TEST()
 * Minimal API call to confirm Gemini connectivity.
 */
export async function TEST() {
  try {
    const res = await geminiMinimalTest({ timeoutMs: DEFAULTS?.timeoutMs || 15000 });
    if (!res.ok) return errorCode(res.code);
    return "OK";
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.ASK(prompt, [contextRange], [options])
 */
export async function ASK(prompt, contextRange, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";

    const ctx = contextRange ? matrixToTSV(contextRange, opt.maxContextChars || LIMITS.MAX_CONTEXT_CHARS) : "";
    const user = [ctx ? `CONTEXT (TSV, may be truncated):\n${ctx}` : "", `USER PROMPT:\n${coerceToTextOrJoin2D(prompt)}`]
      .filter(Boolean)
      .join("\n\n");

    const res = await callGemini({
      system: sysAsk(lang),
      user,
      options: opt
    });

    if (!res.ok) return errorCode(res.code);
    return truncateForCell(res.text);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.TRANSLATE(text, targetLang, [options])
 */
export async function TRANSLATE(text, targetLang, options) {
  try {
    const opt = parseOptions(options);
    const lang = safeString(targetLang).trim() || "en";
    const res = await callGemini({
      system: sysTranslate(lang),
      user: normalizeNewlines(coerceToTextOrJoin2D(text)),
      options: opt
    });
    if (!res.ok) return errorCode(res.code);
    return truncateForCell(res.text);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.CLASSIFY(text, labels, [options])
 */
export async function CLASSIFY(text, labels, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "en";
    const threshold = typeof opt.threshold === "number" ? clamp(opt.threshold, 0, 1, 0.55) : 0.55;

    const labs = flattenLabels(labels);
    if (!labs.length) return errorCode(ERR.BAD_INPUT);

    const system = sysClassify(labs, lang);
    const user = [
      "TEXT:",
      normalizeNewlines(coerceToTextOrJoin2D(text)),
      "",
      `Return only one label. If confidence < ${threshold}, return UNKNOWN.`
    ].join("\n");

    const res = await callGemini({ system, user, options: opt });
    if (!res.ok) return errorCode(res.code);

    const out = normalizeNewlines(res.text).trim();
    // strict output
    const outUpper = out.toUpperCase();
    if (outUpper === "UNKNOWN") return "UNKNOWN";

    // match one of labels (case-insensitive)
    const match = labs.find((l) => l.toLowerCase() === out.toLowerCase());
    return match || "UNKNOWN";
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.CLEAN(text, [options])
 */
export async function CLEAN(text, options) {
  try {
    const opt = parseOptions(options);
    const mode = (opt.mode || "basic").toLowerCase();

    const raw = normalizeNewlines(coerceToTextOrJoin2D(text));
    if (!raw.trim()) return "";

    if (mode === "basic") {
      // non-AI deterministic
      let s = raw.trim();
      s = s.replace(/[ \t]+/g, " ");
      s = s.replace(/\n{3,}/g, "\n\n");
      if (opt.case === "upper") s = s.toUpperCase();
      if (opt.case === "lower") s = s.toLowerCase();
      return truncateForCell(s);
    }

    const lang = opt.lang || "fr";
    const res = await callGemini({
      system: sysClean(lang),
      user: raw,
      options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0 }
    });

    if (!res.ok) return errorCode(res.code);
    return truncateForCell(res.text);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.SUMMARIZE(textOrRange, [options])
 */
export async function SUMMARIZE(textOrRange, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";
    const raw = normalizeNewlines(coerceToTextOrJoin2D(textOrRange));
    if (!raw.trim()) return "";

    const res = await callGemini({
      system: sysSummarize(lang),
      user: raw,
      options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.2 }
    });

    if (!res.ok) return errorCode(res.code);
    return truncateForCell(res.text);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

// --- EXTRACT helpers ---

function parseSchemaOrFields(schemaOrFields) {
  const raw = safeString(schemaOrFields).trim();
  if (!raw) return { ok: false, fields: [], types: {} };

  // try JSON schema
  if (raw.startsWith("{") && raw.endsWith("}")) {
    const p = safeJsonParse(raw);
    if (p.ok && p.value && typeof p.value === "object" && !Array.isArray(p.value)) {
      const fields = Object.keys(p.value).map((k) => k.trim()).filter(Boolean);
      const types = {};
      for (const k of fields) {
        const t = p.value[k];
        types[k] = typeof t === "string" ? t : safeString(t);
      }
      return { ok: fields.length > 0, fields, types };
    }
  }

  // treat as list
  const fields = raw
    .split(/[,;\n]+/)
    .map((s) => s.trim())
    .filter(Boolean);

  return { ok: fields.length > 0, fields, types: {} };
}

function extractJsonObject(text) {
  const s = String(text || "").trim();
  if (!s) return null;

  // If already valid JSON object
  const p1 = safeJsonParse(s);
  if (p1.ok && p1.value && typeof p1.value === "object") return p1.value;

  // Try to locate first {...} block
  const start = s.indexOf("{");
  const end = s.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const candidate = s.slice(start, end + 1);
    const p2 = safeJsonParse(candidate);
    if (p2.ok && p2.value && typeof p2.value === "object") return p2.value;
  }
  return null;
}

/**
 * =AI.EXTRACT(text, schemaOrFields, [options])
 * returns spill: [field | value] unless options.return="json"
 */
export async function EXTRACT(text, schemaOrFields, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";
    const schema = parseSchemaOrFields(schemaOrFields);
    if (!schema.ok) return errorCode(ERR.BAD_SCHEMA);

    const raw = normalizeNewlines(coerceToTextOrJoin2D(text));
    if (!raw.trim()) {
      if (opt.return === "json") return "{}";
      return schema.fields.map((f) => [f, ""]);
    }

    const res = await callGemini({
      system: sysExtract(schema.fields, lang),
      user: raw,
      options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0 }
    });

    if (!res.ok) return errorCode(res.code);

    const obj = extractJsonObject(res.text);
    if (!obj) return errorCode(ERR.PARSE_ERROR);

    if (opt.return === "json") {
      // return compact JSON string
      try {
        return truncateForCell(JSON.stringify(obj));
      } catch {
        return errorCode(ERR.PARSE_ERROR);
      }
    }

    // 2-col spill
    const out = [];
    for (const f of schema.fields) {
      const v = obj[f];
      out.push([f, v === null || v === undefined ? "" : safeString(v)]);
    }
    return out;
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.TABLE(prompt, [contextRange], [options])
 * returns 2D table with headers in first row
 */
export async function TABLE(prompt, contextRange, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";
    const maxRows = clamp(opt.maxRows, 1, LIMITS.MAX_TABLE_ROWS, LIMITS.MAX_TABLE_ROWS);
    const maxCols = clamp(opt.maxCols, 1, LIMITS.MAX_TABLE_COLS, LIMITS.MAX_TABLE_COLS);
    const headers = Array.isArray(opt.headers) ? opt.headers.map((h) => safeString(h)) : null;

    const ctx = contextRange ? matrixToTSV(contextRange, opt.maxContextChars || LIMITS.MAX_CONTEXT_CHARS) : "";
    const user = [ctx ? `CONTEXT (TSV, may be truncated):\n${ctx}` : "", `PROMPT:\n${coerceToTextOrJoin2D(prompt)}`]
      .filter(Boolean)
      .join("\n\n");

    const res = await callGemini({
      system: sysTable(lang, maxRows, maxCols, headers),
      user,
      options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.1 }
    });

    if (!res.ok) return errorCode(res.code);

    const obj = extractJsonObject(res.text);
    if (!obj) return errorCode(ERR.PARSE_ERROR);

    const h = Array.isArray(obj.headers) ? obj.headers.map((x) => safeString(x)) : [];
    const rows = Array.isArray(obj.rows) ? obj.rows : [];
    if (!h.length) return errorCode(ERR.PARSE_ERROR);

    const clippedHeaders = h.slice(0, maxCols);
    const out = [clippedHeaders];

    const rowCount = Math.min(rows.length, maxRows);
    for (let i = 0; i < rowCount; i++) {
      const r = Array.isArray(rows[i]) ? rows[i] : [];
      const line = [];
      for (let c = 0; c < clippedHeaders.length; c++) {
        const v = r[c];
        line.push(v === null || v === undefined ? "" : safeString(v));
      }
      out.push(line);
    }

    return out;
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.FILL(exampleRange, targetRange, instruction, [options])
 * Returns spill values corresponding to target rows.
 */
export async function FILL(exampleRange, targetRange, instruction, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";

    const examples = Array.isArray(exampleRange) ? exampleRange : [];
    const targets = Array.isArray(targetRange) ? targetRange : [];

    // Derive number of rows to fill from targetRange
    const targetRows = Array.isArray(targets) ? targets.length : 0;
    if (!targetRows) return [];

    const maxRows = clamp(opt.maxRows, 1, LIMITS.MAX_FILL_ROWS, Math.min(LIMITS.MAX_FILL_ROWS, targetRows));
    const rowsToFill = Math.min(targetRows, maxRows);

    // Build compact example pairs: take first 40 rows of [input | output]
    const exTSV = matrixToTSV(examples, 2000);
    const tgtTSV = matrixToTSV(targets.slice(0, rowsToFill), 4000);

    const user = [
      "INSTRUCTION:",
      normalizeNewlines(coerceToTextOrJoin2D(instruction)),
      "",
      "EXAMPLES (TSV):",
      exTSV,
      "",
      "TARGET INPUTS (TSV):",
      tgtTSV
    ].join("\n");

    const res = await callGemini({
      system: sysFill(lang, rowsToFill),
      user,
      options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0 }
    });

    if (!res.ok) return errorCode(res.code);

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.values)) return errorCode(ERR.PARSE_ERROR);

    const values = obj.values.slice(0, rowsToFill).map((x) => safeString(x));
    // spill as column
    return values.map((v) => [truncateForCell(v, LIMITS.MAX_CELL_CHARS)]);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * Register custom functions by associating the JSON metadata `id` values to implementations.
 *
 * IMPORTANT:
 * - In functions.json, each `id` can only contain alphanumeric characters and periods.
 *   That is why the status function uses id "AI.KEYSTATUS" even though the Excel function
 *   name exposed to the user is AI.KEY_STATUS (because the *name* in functions.json is KEY_STATUS).
 * - Never let a single failed association prevent registering the other functions.
 */
function registerCustomFunctions() {
  if (typeof CustomFunctions === "undefined" || typeof CustomFunctions.associate !== "function") return false;

  const pairs = [
    ["AI.ASK", ASK],
    ["AI.EXTRACT", EXTRACT],
    ["AI.CLASSIFY", CLASSIFY],
    ["AI.TRANSLATE", TRANSLATE],
    ["AI.TABLE", TABLE],
    ["AI.FILL", FILL],
    ["AI.CLEAN", CLEAN],
    ["AI.SUMMARIZE", SUMMARIZE],
    ["AI.KEYSTATUS", KEY_STATUS],
    ["AI.TEST", TEST]
  ];

  let any = false;
  for (const [id, fn] of pairs) {
    try {
      CustomFunctions.associate(id, fn);
      any = true;
    } catch (e) {
      // Keep going so one bad association doesn't break all functions.
      try { console.warn(`[AI] CustomFunctions.associate failed for ${id}`, e); } catch { /* ignore */ }
    }
  }

  return any;
}

// Attempt immediate registration.
const _registered = registerCustomFunctions();

// If CustomFunctions wasn't ready yet, retry a few times (non-blocking).
if (!_registered && typeof setTimeout === "function") {
  let attempts = 0;
  const maxAttempts = 20; // ~10s
  const intervalMs = 500;

  const tick = () => {
    attempts++;
    if (registerCustomFunctions() || attempts >= maxAttempts) return;
    setTimeout(tick, intervalMs);
  };

  setTimeout(tick, 0);
}
