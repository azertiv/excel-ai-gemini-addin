// src/functions/functions.js
/* global CustomFunctions */

import { geminiGenerate, geminiMinimalTest } from "../shared/gemini.js";
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
    return { _invalidOptions: true };
  }
}

function errorCode(code) {
  // Prefer your ERR constants; fall back safely
  if (typeof code === "string" && code.startsWith("#")) return code;
  return ERR?.API_ERROR || "#AI_API_ERROR";
}

function to2DRange(values) {
  // values can be string "a|b|c" or 2D array from Excel.
  if (Array.isArray(values)) {
    // flatten 2D into list of non-empty strings
    const out = [];
    for (const row of values) {
      if (!Array.isArray(row)) continue;
      for (const cell of row) {
        const s = safeString(cell).trim();
        if (s) out.push(s);
      }
    }
    return out;
  }
  // string form: "A|B|C"
  const s = safeString(values).trim();
  if (!s) return [];
  return s.split("|").map((x) => x.trim()).filter(Boolean);
}

function contextToText(contextRange, maxChars = 3500) {
  if (!contextRange) return "";
  if (!Array.isArray(contextRange)) return "";

  // Serialize as TSV-like lines, trimmed
  const lines = [];
  for (const row of contextRange) {
    if (!Array.isArray(row)) continue;
    lines.push(row.map((c) => safeString(c)).join("\t"));
    if (lines.join("\n").length > maxChars) break;
  }
  let t = lines.join("\n");
  if (t.length > maxChars) t = t.slice(0, maxChars) + "\n…(truncated)";
  return t;
}

function clamp(n, lo, hi) {
  n = Number(n);
  if (!Number.isFinite(n)) return lo;
  return Math.max(lo, Math.min(hi, n));
}

function buildGenConfig(opt) {
  const temperature = typeof opt.temperature === "number" ? opt.temperature : (DEFAULTS?.temperature ?? 0.2);
  const maxOutputTokens = typeof opt.maxTokens === "number" ? opt.maxTokens : (DEFAULTS?.maxTokens ?? 256);
  return {
    temperature: clamp(temperature, 0, 1),
    maxOutputTokens: clamp(maxOutputTokens, 1, 4096),
  };
}

function maxCellTrim(text) {
  let t = safeString(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  t = t.trim();
  const limit = LIMITS?.MAX_CELL_CHARS || 28000;
  if (t.length > limit) t = t.slice(0, limit) + "\n…(truncated)";
  return t;
}

async function callText({ system, user, optionsJson, cacheMode }) {
  const opt = parseOptions(optionsJson);
  if (opt._invalidOptions) return { ok: false, code: ERR.BAD_INPUT || "#AI_BAD_INPUT", message: "Invalid options JSON" };

  const res = await geminiGenerate({
    model: opt.model,
    system,
    user,
    generationConfig: buildGenConfig(opt),
    cache: opt.cache || cacheMode || DEFAULTS?.cache || "memory",
    cacheTtlSec: typeof opt.cacheTtlSec === "number" ? opt.cacheTtlSec : DEFAULTS?.cacheTtlSec,
    timeoutMs: typeof opt.timeoutMs === "number" ? opt.timeoutMs : DEFAULTS?.timeoutMs,
    retry: typeof opt.retry === "number" ? opt.retry : DEFAULTS?.retry,
    // responseMimeType/responseJsonSchema supported in your gemini.js already if needed
  });

  return res;
}

function parseJsonStrict(s) {
  const txt = safeString(s).trim();
  if (!txt) return null;
  // remove fenced code blocks if any
  const unfenced = txt
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/```$/i, "")
    .trim();
  try {
    return JSON.parse(unfenced);
  } catch {
    return null;
  }
}

function fieldsFromSchema(schemaOrFields) {
  const raw = safeString(schemaOrFields).trim();
  if (!raw) return [];
  // Try JSON schema map: {"email":"string",...}
  const maybeObj = parseJsonStrict(raw);
  if (maybeObj && typeof maybeObj === "object" && !Array.isArray(maybeObj)) {
    return Object.keys(maybeObj);
  }
  // Comma list: "email, phone, amount"
  return raw.split(",").map((x) => x.trim()).filter(Boolean);
}

function objectTo2ColSpill(obj) {
  const rows = [["field", "value"]];
  if (obj && typeof obj === "object") {
    for (const k of Object.keys(obj)) {
      const v = obj[k];
      rows.push([k, (v === null || v === undefined) ? "" : (typeof v === "string" ? v : JSON.stringify(v))]);
    }
  }
  return rows;
}

// ---------- Custom Functions (public) ----------

/**
 * =AI.KEY_STATUS()
 */
export async function KEY_STATUS() {
  try {
    // geminiGenerate will return KEY_MISSING if not present; but for status we can shortcut by calling test with no key.
    // We do not have direct access to key here (stored in shared/storage). So use geminiGenerate semantics:
    // Instead, rely on geminiMinimalTest which will return KEY_MISSING without key.
    const res = await geminiMinimalTest({ timeoutMs: 2000 });
    if (!res.ok && (res.code === ERR.KEY_MISSING || res.code === "#AI_KEY_MISSING")) return "MISSING";
    return "OK";
  } catch {
    return "MISSING";
  }
}

/**
 * =AI.TEST()
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
    const p = safeString(prompt).trim();
    if (!p) return errorCode(ERR.BAD_INPUT);

    const opt = parseOptions(options);
    if (opt._invalidOptions) return errorCode(ERR.BAD_INPUT);

    const lang = safeString(opt.lang || "fr").trim();
    const ctx = contextToText(contextRange, 3500);

    const system =
      `You are an assistant for Excel users. Respond concisely (1-10 lines).` +
      ` If the user requests a language, comply. Output plain text only.`;

    const user =
      (ctx ? `Context (table/range):\n${ctx}\n\n` : "") +
      `User request (language=${lang}):\n${p}`;

    const res = await callText({ system, user, optionsJson: options, cacheMode: "memory" });
    if (!res.ok) return errorCode(res.code);
    return maxCellTrim(res.text);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.TRANSLATE(text, targetLang, [options])
 */
export async function TRANSLATE(text, targetLang, options) {
  try {
    const t = safeString(text).trim();
    if (!t) return "";

    const lang = safeString(targetLang).trim();
    if (!lang) return errorCode(ERR.BAD_INPUT);

    const system = `You translate text. Output ONLY the translated text. No quotes.`;
    const user = `Translate to ${lang}:\n${t}`;

    const res = await callText({ system, user, optionsJson: options, cacheMode: "memory" });
    if (!res.ok) return errorCode(res.code);
    return maxCellTrim(res.text);
  } catch {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.CLASSIFY(text, labels, [options])
 * labels can be "A|B|C" or a range
 */
export async function CLASSIFY(text, labels, options) {
  try {
    const t = safeString(text).trim();
    if (!t) return "UNKNOWN";

    const opt = parseOptions(options);
    if (opt._invalidOptions) return errorCode(ERR.BAD_INPUT);

    const threshold = typeof opt.threshold === "number" ? opt.threshold : 0.6;
    const labs = Array.isArray(labels) ? to2DRange(labels) : to2DRange(safeString(labels));
    if (!labs.length) return errorCode(ERR.BAD_INPUT);

    const system =
      `You are a classifier. Choose exactly ONE label from the provided list.` +
      ` If unsure, output EXACTLY "UNKNOWN". Output one token / one label only.`;

    const user =
      `Labels: ${labs.join(" | ")}\n\n` +
      `Text:\n${t}\n\n` +
      `Return one label or UNKNOWN.`;

    // encourage determinism
    const localOptions = JSON.stringify({ ...(opt || {}), temperature: 0.0, maxTokens: Math.max(16, opt.maxTokens || 16) });

    const res = await callText({ system, user, optionsJson: localOptions, cacheMode: "memory" });
    if (!res.ok) return errorCode(res.code);

    const out = safeString(res.text).trim();
    if (!out) return "UNKNOWN";

    // If model returns something not in labels, map to UNKNOWN
    const normalized = out.replace(/["'.]/g, "").trim();
    const hit = labs.find((l) => l.toLowerCase() === normalized.toLowerCase());
    if (hit) return hit;

    // If user sets threshold, we can't compute real probability; keep UNKNOWN as safe default.
    if (threshold >= 0.0) return "UNKNOWN";
    return "UNKNOWN";
  } catch {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.EXTRACT(text, schemaOrFields, [options])
 * Returns spill 2 columns: field | value OR JSON if options.return="json"
 */
export async function EXTRACT(text, schemaOrFields, options) {
  try {
    const input = safeString(text).trim();
    if (!input) return [["field", "value"]];

    const opt = parseOptions(options);
    if (opt._invalidOptions) return [[errorCode(ERR.BAD_INPUT), "Invalid options JSON"]];

    const fields = fieldsFromSchema(schemaOrFields);
    if (!fields.length) return [[errorCode(ERR.BAD_INPUT), "Missing fields/schema"]];

    const wantJson = safeString(opt.return || "").toLowerCase() === "json";

    const system =
      `Extract structured fields from text.` +
      ` Output STRICT JSON object with keys exactly as requested. Do not include extra keys.` +
      ` Use empty string for missing values.`;

    const user =
      `Fields: ${fields.join(", ")}\n\n` +
      `Text:\n${input}\n\n` +
      `Return JSON only.`;

    const localOptions = JSON.stringify({ ...(opt || {}), temperature: 0.0, maxTokens: Math.max(256, opt.maxTokens || 256) });

    const res = await callText({ system, user, optionsJson: localOptions, cacheMode: "memory" });
    if (!res.ok) return [[errorCode(res.code), res.message || "Error"]];

    const obj = parseJsonStrict(res.text);
    if (!obj || typeof obj !== "object") {
      // fallback: return raw text to help debugging
      return [[errorCode(ERR.API_ERROR), "Bad JSON output"]];
    }

    if (wantJson) return maxCellTrim(JSON.stringify(obj));
    return objectTo2ColSpill(obj);
  } catch {
    return [[errorCode(ERR.API_ERROR), "Error"]];
  }
}

/**
 * =AI.TABLE(prompt, [contextRange], [options])
 * Returns 2D array (spill), first row headers
 */
export async function TABLE(prompt, contextRange, options) {
  try {
    const p = safeString(prompt).trim();
    if (!p) return [[errorCode(ERR.BAD_INPUT), "Missing prompt"]];

    const opt = parseOptions(options);
    if (opt._invalidOptions) return [[errorCode(ERR.BAD_INPUT), "Invalid options JSON"]];

    const ctx = contextToText(contextRange, 3500);
    const maxRows = typeof opt.maxRows === "number" ? Math.max(1, Math.min(200, Math.floor(opt.maxRows))) : 20;
    const headers = Array.isArray(opt.headers) ? opt.headers.map(safeString) : [];
    const nCols = typeof opt.nbColonnes === "number" ? Math.max(1, Math.min(30, Math.floor(opt.nbColonnes))) : 0;

    const system =
      `You generate tabular data for Excel.` +
      ` Output STRICT JSON with shape: {"headers":[...], "rows":[[...],[...]]}.` +
      ` No markdown. No prose.`;

    const user =
      (ctx ? `Context:\n${ctx}\n\n` : "") +
      `Task:\n${p}\n\n` +
      (headers.length ? `Preferred headers: ${headers.join(", ")}\n` : "") +
      (nCols ? `Number of columns: ${nCols}\n` : "") +
      `Max rows: ${maxRows}\n\n` +
      `Return JSON only.`;

    const localOptions = JSON.stringify({ ...(opt || {}), temperature: 0.0, maxTokens: Math.max(512, opt.maxTokens || 512) });

    const res = await callText({ system, user, optionsJson: localOptions, cacheMode: "memory" });
    if (!res.ok) return [[errorCode(res.code), res.message || "Error"]];

    const obj = parseJsonStrict(res.text);
    if (!obj || typeof obj !== "object") return [[errorCode(ERR.API_ERROR), "Bad JSON output"]];

    const outHeaders = Array.isArray(obj.headers) ? obj.headers.map(safeString) : [];
    const outRows = Array.isArray(obj.rows) ? obj.rows : [];

    if (!outHeaders.length || !outRows.length) {
      return [[errorCode(ERR.EMPTY_RESPONSE), "Empty table"]];
    }

    // Normalize rows length to headers length
    const hLen = outHeaders.length;
    const rows2d = [outHeaders];
    for (let i = 0; i < Math.min(outRows.length, maxRows); i++) {
      const r = Array.isArray(outRows[i]) ? outRows[i] : [];
      const row = [];
      for (let c = 0; c < hLen; c++) row.push(safeString(r[c] ?? ""));
      rows2d.push(row);
    }
    return rows2d;
  } catch {
    return [[errorCode(ERR.API_ERROR), "Error"]];
  }
}

/**
 * =AI.FILL(exampleRange, targetRange, instruction, [options])
 * Returns a spill corresponding to suggested values for targetRange rows.
 */
export async function FILL(exampleRange, targetRange, instruction, options) {
  try {
    const instr = safeString(instruction).trim();
    if (!instr) return [[errorCode(ERR.BAD_INPUT), "Missing instruction"]];

    const opt = parseOptions(options);
    if (opt._invalidOptions) return [[errorCode(ERR.BAD_INPUT), "Invalid options JSON"]];

    const ex = contextToText(exampleRange, 3000);
    const tgt = contextToText(targetRange, 2000);

    const system =
      `You are helping fill Excel columns.` +
      ` Return STRICT JSON: {"rows":["v1","v2",...]} with one value per target row.` +
      ` No markdown.`;

    const user =
      `Instruction:\n${instr}\n\n` +
      `Examples (range):\n${ex}\n\n` +
      `Target (range):\n${tgt}\n\n` +
      `Return JSON only.`;

    const localOptions = JSON.stringify({ ...(opt || {}), temperature: 0.0, maxTokens: Math.max(512, opt.maxTokens || 512) });

    const res = await callText({ system, user, optionsJson: localOptions, cacheMode: "memory" });
    if (!res.ok) return [[errorCode(res.code), res.message || "Error"]];

    const obj = parseJsonStrict(res.text);
    const rows = Array.isArray(obj?.rows) ? obj.rows : null;
    if (!rows) return [[errorCode(ERR.API_ERROR), "Bad JSON output"]];

    // Spill as single column
    return rows.slice(0, 500).map((v) => [safeString(v)]);
  } catch {
    return [[errorCode(ERR.API_ERROR), "Error"]];
  }
}

/**
 * =AI.CLEAN(text, [options])
 */
export async function CLEAN(text, options) {
  try {
    const t = safeString(text);
    if (!t.trim()) return "";

    const opt = parseOptions(options);
    if (opt._invalidOptions) return errorCode(ERR.BAD_INPUT);

    // Non-AI clean by default
    const mode = safeString(opt.mode || "basic").toLowerCase();
    if (mode === "basic") {
      return t
        .replace(/\s+/g, " ")
        .trim();
    }

    // AI clean if requested
    const system = `Normalize the text. Output ONLY the cleaned text.`;
    const user = `Clean/normalize:\n${t}`;
    const localOptions = JSON.stringify({ ...(opt || {}), temperature: 0.0, maxTokens: Math.max(128, opt.maxTokens || 128) });

    const res = await callText({ system, user, optionsJson: localOptions, cacheMode: "memory" });
    if (!res.ok) return errorCode(res.code);
    return maxCellTrim(res.text);
  } catch {
    return errorCode(ERR.API_ERROR);
  }
}

/**
 * =AI.SUMMARIZE(textOrRange, [options])
 */
export async function SUMMARIZE(textOrRange, options) {
  try {
    const opt = parseOptions(options);
    if (opt._invalidOptions) return errorCode(ERR.BAD_INPUT);

    let t = "";
    if (Array.isArray(textOrRange)) {
      t = contextToText(textOrRange, 3500);
    } else {
      t = safeString(textOrRange);
    }
    t = t.trim();
    if (!t) return "";

    const system = `Summarize content into concise bullet points. Output plain text bullets only.`;
    const user = `Summarize:\n${t}`;

    const localOptions = JSON.stringify({ ...(opt || {}), temperature: 0.2, maxTokens: Math.max(256, opt.maxTokens || 256) });

    const res = await callText({ system, user, optionsJson: localOptions, cacheMode: "memory" });
    if (!res.ok) return errorCode(res.code);
    return maxCellTrim(res.text);
  } catch {
    return errorCode(ERR.API_ERROR);
  }
}

// ---------- Associations (important) ----------
// This prevents “function exists but not bound” issues in some environments.
try {
  // Support both the correct id (AI.KEY_STATUS) and the legacy id (AI.KEYSTATUS)
  CustomFunctions.associate("AI.KEY_STATUS", KEY_STATUS);
  CustomFunctions.associate("AI.KEYSTATUS", KEY_STATUS);
  CustomFunctions.associate("AI.TEST", TEST);
  CustomFunctions.associate("AI.ASK", ASK);
  CustomFunctions.associate("AI.EXTRACT", EXTRACT);
  CustomFunctions.associate("AI.CLASSIFY", CLASSIFY);
  CustomFunctions.associate("AI.TRANSLATE", TRANSLATE);
  CustomFunctions.associate("AI.TABLE", TABLE);
  CustomFunctions.associate("AI.FILL", FILL);
  CustomFunctions.associate("AI.CLEAN", CLEAN);
  CustomFunctions.associate("AI.SUMMARIZE", SUMMARIZE);
} catch {
  // ignore (Excel will still try to bind via metadata, but associate is recommended)
}
