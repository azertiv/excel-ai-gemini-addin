import { ERR, LIMITS } from "../shared/constants";
import { parseOptions, normalizedCommonOptions, clamp } from "../shared/options";
import { isMatrix, normalizeTextInput, matrixToTSV, flattenToStringList } from "../shared/range";
import { geminiGenerate, geminiMinimalTest } from "../shared/gemini";
import {
  buildAskPrompt,
  parseFields,
  buildExtractPrompt,
  buildClassifyPrompt,
  buildTranslatePrompt,
  buildTablePrompt,
  buildFillPrompt,
  buildSummarizePrompt,
  buildCleanAiPrompt
} from "../shared/prompts";
import { getApiKey } from "../shared/storage";

function looksLikeJsonObjectString(s) {
  if (typeof s !== "string") return false;
  const t = s.trim();
  return t.startsWith("{") && t.endsWith("}");
}

function safeJsonParse(text) {
  if (typeof text !== "string") return { ok: false, error: "not_string" };
  const t = text.trim();
  if (!t) return { ok: false, error: "empty" };
  try {
    return { ok: true, value: JSON.parse(t) };
  } catch (e) {
    const firstObj = t.indexOf("{");
    const lastObj = t.lastIndexOf("}");
    if (firstObj !== -1 && lastObj !== -1 && lastObj > firstObj) {
      const sub = t.slice(firstObj, lastObj + 1);
      try { return { ok: true, value: JSON.parse(sub) }; } catch { /* ignore */ }
    }
    const firstArr = t.indexOf("[");
    const lastArr = t.lastIndexOf("]");
    if (firstArr !== -1 && lastArr !== -1 && lastArr > firstArr) {
      const sub = t.slice(firstArr, lastArr + 1);
      try { return { ok: true, value: JSON.parse(sub) }; } catch { /* ignore */ }
    }
    return { ok: false, error: e?.message || "parse_error" };
  }
}

function to2D(value) {
  if (Array.isArray(value) && Array.isArray(value[0])) return value;
  return [[value]];
}

function stringifyValue(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  try { return JSON.stringify(v); } catch { return String(v); }
}

function normalizeOneCellText(s) {
  let out = (s ?? "").toString();
  out = out.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
  if (out.length > LIMITS.MAX_CELL_CHARS) out = out.slice(0, LIMITS.MAX_CELL_CHARS) + "\n…(truncated)";
  return out;
}

function parseLabels(labels) {
  if (isMatrix(labels)) {
    const list = flattenToStringList(labels);
    return Array.from(new Set(list));
  }
  const s = (labels ?? "").toString().trim();
  if (!s) return [];
  const list = s.split(/[|,\n;]+/).map((x) => x.trim()).filter(Boolean);
  return Array.from(new Set(list));
}

function parseHeaders(headersOpt) {
  if (!headersOpt) return [];
  if (Array.isArray(headersOpt)) return headersOpt.map((h) => String(h).trim()).filter(Boolean);
  const s = String(headersOpt).trim();
  if (!s) return [];
  return s.split(/[|,\n;]+/).map((x) => x.trim()).filter(Boolean);
}

/** =AI.ASK(prompt, [contextRange], [options]) */
export async function ASK(prompt, contextRange, options) {
  if (typeof contextRange === "string" && options === undefined && looksLikeJsonObjectString(contextRange)) {
    options = contextRange;
    contextRange = undefined;
  }

  const p = normalizeTextInput(prompt, 4000).trim();
  if (!p) return ERR.BAD_INPUT;

  const parsed = parseOptions(options);
  if (!parsed.ok) return ERR.BAD_OPTIONS;

  const opts = normalizedCommonOptions(parsed.value);
  const { system, user } = buildAskPrompt(p, contextRange, { ...opts, maxContextChars: LIMITS.MAX_CONTEXT_CHARS });

  const res = await geminiGenerate({
    model: opts.model,
    system,
    user,
    generationConfig: {
      temperature: clamp(opts.temperature, 0, 1, 0.2),
      maxOutputTokens: clamp(opts.maxTokens, 1, 2048, 256)
    },
    cache: opts.cache,
    cacheTtlSec: opts.cacheTtlSec,
    timeoutMs: opts.timeoutMs,
    retry: opts.retry
  });

  if (!res.ok) return res.code;
  return normalizeOneCellText(res.text);
}

/** =AI.EXTRACT(text, schemaOrFields, [options]) */
export async function EXTRACT(text, schemaOrFields, options) {
  const parsedOptions = parseOptions(options);
  if (!parsedOptions.ok) return to2D(ERR.BAD_OPTIONS);

  const userOpt = parsedOptions.value || {};
  const opts = normalizedCommonOptions(userOpt);
  const returnMode = typeof userOpt.return === "string" ? userOpt.return.toLowerCase() : "table";

  const schemaInfo = parseFields(schemaOrFields);
  if (!schemaInfo.ok || !schemaInfo.fields?.length) return to2D(ERR.BAD_SCHEMA);

  const fields = schemaInfo.fields.slice(0, 50);
  const inputText = normalizeTextInput(text, LIMITS.MAX_INPUT_CHARS).trim();

  if (!inputText) {
    if (returnMode === "json") {
      const obj = {};
      for (const f of fields) obj[f] = "";
      return to2D(JSON.stringify(obj));
    }
    return fields.map((f) => [f, ""]);
  }

  const { system, user } = buildExtractPrompt(inputText, schemaInfo, opts);

  const properties = {};
  for (const f of fields) properties[f] = { type: "string" };

  const responseJsonSchema = {
    type: "object",
    properties,
    required: fields,
    additionalProperties: false
  };

  const res = await geminiGenerate({
    model: opts.model,
    system,
    user,
    generationConfig: { temperature: 0.0, maxOutputTokens: clamp(userOpt.maxTokens ?? opts.maxTokens ?? 512, 32, 2048, 512) },
    responseMimeType: "application/json",
    responseJsonSchema,
    cache: opts.cache,
    cacheTtlSec: opts.cacheTtlSec,
    timeoutMs: opts.timeoutMs,
    retry: opts.retry
  });

  if (!res.ok) return to2D(res.code);

  const parsed = safeJsonParse(res.text);
  if (!parsed.ok || !parsed.value || typeof parsed.value !== "object") return to2D(ERR.PARSE_ERROR);

  const obj = parsed.value;

  if (returnMode === "json") {
    const out = {};
    for (const f of fields) out[f] = stringifyValue(obj[f] ?? "");
    return to2D(JSON.stringify(out));
  }

  return fields.map((f) => [f, stringifyValue(obj[f] ?? "")]);
}

/** =AI.CLASSIFY(text, labels, [options]) */
export async function CLASSIFY(text, labels, options) {
  const parsedOptions = parseOptions(options);
  if (!parsedOptions.ok) return ERR.BAD_OPTIONS;

  const userOpt = parsedOptions.value || {};
  const opts = normalizedCommonOptions(userOpt);

  const labelList = parseLabels(labels);
  if (!labelList.length) return ERR.BAD_INPUT;

  const unknownLabel = typeof userOpt.unknownLabel === "string" ? userOpt.unknownLabel : "UNKNOWN";
  const threshold = clamp(userOpt.threshold, 0, 1, 0.65);

  const inputText = normalizeTextInput(text, 8000).trim();
  if (!inputText) return unknownLabel;

  const allowed = Array.from(new Set([...labelList, unknownLabel]));
  const responseJsonSchema = {
    type: "object",
    properties: {
      label: { type: "string", enum: allowed },
      confidence: { type: "number" }
    },
    required: ["label", "confidence"],
    additionalProperties: false
  };

  const { system, user } = buildClassifyPrompt(inputText, labelList, { ...opts, unknownLabel });

  const res = await geminiGenerate({
    model: opts.model,
    system,
    user,
    generationConfig: { temperature: 0.0, maxOutputTokens: 64 },
    responseMimeType: "application/json",
    responseJsonSchema,
    cache: opts.cache,
    cacheTtlSec: opts.cacheTtlSec,
    timeoutMs: opts.timeoutMs,
    retry: opts.retry
  });

  if (!res.ok) return res.code;

  const parsed = safeJsonParse(res.text);
  if (!parsed.ok || !parsed.value) return ERR.PARSE_ERROR;

  const label = (parsed.value.label ?? "").toString().trim();
  const conf = Number(parsed.value.confidence);

  if (!allowed.includes(label)) return unknownLabel;
  if (!Number.isFinite(conf) || conf < threshold) return unknownLabel;
  return label;
}

/** =AI.TRANSLATE(text, targetLang, [options]) */
export async function TRANSLATE(text, targetLang, options) {
  const parsed = parseOptions(options);
  if (!parsed.ok) return ERR.BAD_OPTIONS;

  const opts = normalizedCommonOptions(parsed.value);
  const t = normalizeTextInput(text, LIMITS.MAX_INPUT_CHARS).trim();
  if (!t) return "";

  const lang = (targetLang ?? "").toString().trim();
  if (!lang) return ERR.BAD_INPUT;

  const { system, user } = buildTranslatePrompt(t, lang, opts);

  const res = await geminiGenerate({
    model: opts.model,
    system,
    user,
    generationConfig: { temperature: clamp(opts.temperature, 0, 1, 0.2), maxOutputTokens: clamp(opts.maxTokens, 1, 4096, 512) },
    cache: opts.cache,
    cacheTtlSec: opts.cacheTtlSec,
    timeoutMs: opts.timeoutMs,
    retry: opts.retry
  });

  if (!res.ok) return res.code;
  return normalizeOneCellText(res.text);
}

/** =AI.TABLE(prompt, [contextRange], [options]) */
export async function TABLE(prompt, contextRange, options) {
  if (typeof contextRange === "string" && options === undefined && looksLikeJsonObjectString(contextRange)) {
    options = contextRange;
    contextRange = undefined;
  }

  const p = normalizeTextInput(prompt, 4000).trim();
  if (!p) return to2D(ERR.BAD_INPUT);

  const parsedOptions = parseOptions(options);
  if (!parsedOptions.ok) return to2D(ERR.BAD_OPTIONS);

  const userOpt = parsedOptions.value || {};
  const opts = normalizedCommonOptions(userOpt);

  const headers = parseHeaders(userOpt.headers);
  const numColumns = headers.length ? headers.length : clamp(userOpt.numColumns ?? userOpt.nbColonnes, 1, LIMITS.MAX_TABLE_COLS, 3);
  const maxRows = clamp(userOpt.maxRows ?? userOpt.nbLignesMax ?? userOpt.nbLignes, 1, LIMITS.MAX_TABLE_ROWS, 10);
  const maxCols = clamp(userOpt.maxCols ?? numColumns, 1, LIMITS.MAX_TABLE_COLS, numColumns);

  const { system, user } = buildTablePrompt(p, contextRange, { ...opts, headers, numColumns, maxRows, maxCols });

  const responseJsonSchema = {
    type: "object",
    properties: {
      headers: { type: "array", items: { type: "string" } },
      rows: { type: "array", items: { type: "array", items: { type: "string" } } }
    },
    required: ["headers", "rows"],
    additionalProperties: false
  };

  const res = await geminiGenerate({
    model: opts.model,
    system,
    user,
    generationConfig: { temperature: clamp(opts.temperature, 0, 1, 0.2), maxOutputTokens: clamp(opts.maxTokens ?? 800, 64, 4096, 800) },
    responseMimeType: "application/json",
    responseJsonSchema,
    cache: opts.cache,
    cacheTtlSec: opts.cacheTtlSec,
    timeoutMs: opts.timeoutMs,
    retry: opts.retry
  });

  if (!res.ok) return to2D(res.code);

  const parsed = safeJsonParse(res.text);
  if (!parsed.ok || !parsed.value) return to2D(ERR.PARSE_ERROR);

  let outHeaders = Array.isArray(parsed.value.headers) ? parsed.value.headers.map((h) => String(h).trim()) : [];
  const rows = Array.isArray(parsed.value.rows) ? parsed.value.rows : [];

  if (headers.length) outHeaders = headers.slice(0, LIMITS.MAX_TABLE_COLS);

  if (!outHeaders.length) {
    outHeaders = [];
    for (let i = 0; i < numColumns; i++) outHeaders.push(`Col${i + 1}`);
  }

  outHeaders = outHeaders.slice(0, maxCols);

  const outRows = [];
  for (const r of rows.slice(0, maxRows)) {
    const rr = Array.isArray(r) ? r : [r];
    const line = [];
    for (let c = 0; c < outHeaders.length; c++) line.push(stringifyValue(rr[c] ?? ""));
    outRows.push(line);
  }

  return [outHeaders, ...outRows];
}

/** =AI.FILL(exampleRange, targetRange, instruction, [options]) */
export async function FILL(exampleRange, targetRange, instruction, options) {
  const parsedOptions = parseOptions(options);
  if (!parsedOptions.ok) return to2D(ERR.BAD_OPTIONS);

  const userOpt = parsedOptions.value || {};
  const opts = normalizedCommonOptions(userOpt);

  if (!isMatrix(exampleRange) || !isMatrix(targetRange)) return to2D(ERR.BAD_INPUT);
  if (Array.isArray(targetRange) && targetRange.length > LIMITS.MAX_FILL_ROWS) return to2D(ERR.TOO_LARGE);

  const instr = normalizeTextInput(instruction, 2000).trim();
  if (!instr) return to2D(ERR.BAD_INPUT);

  const { system, user, targetCount } = buildFillPrompt(exampleRange, targetRange, instr, { ...opts, maxExamples: userOpt.maxExamples });
  if (!targetCount) return [[""]];

  const responseJsonSchema = {
    type: "object",
    properties: { values: { type: "array", items: { type: "string" } } },
    required: ["values"],
    additionalProperties: false
  };

  const res = await geminiGenerate({
    model: opts.model,
    system,
    user,
    generationConfig: { temperature: clamp(userOpt.temperature ?? opts.temperature, 0, 1, 0.1), maxOutputTokens: clamp(userOpt.maxTokens ?? opts.maxTokens ?? 1024, 64, 4096, 1024) },
    responseMimeType: "application/json",
    responseJsonSchema,
    cache: opts.cache,
    cacheTtlSec: opts.cacheTtlSec,
    timeoutMs: opts.timeoutMs,
    retry: opts.retry
  });

  if (!res.ok) return to2D(res.code);

  const parsed = safeJsonParse(res.text);
  if (!parsed.ok || !parsed.value) return to2D(ERR.PARSE_ERROR);

  const values = Array.isArray(parsed.value.values) ? parsed.value.values.map((v) => stringifyValue(v)) : [];
  const out = [];
  for (let i = 0; i < targetCount; i++) out.push([values[i] ?? ""]);
  return out;
}

/** =AI.CLEAN(text, [options]) */
export async function CLEAN(text, options) {
  const parsed = parseOptions(options);
  if (!parsed.ok) return ERR.BAD_OPTIONS;

  const userOpt = parsed.value || {};
  const opts = normalizedCommonOptions(userOpt);

  const input = normalizeTextInput(text, LIMITS.MAX_INPUT_CHARS);
  if (!input.trim()) return "";

  const mode = typeof userOpt.mode === "string" ? userOpt.mode.toLowerCase() : (userOpt.ai ? "ai" : "simple");

  if (mode === "ai") {
    const { system, user } = buildCleanAiPrompt(input, opts);
    const res = await geminiGenerate({
      model: opts.model,
      system,
      user,
      generationConfig: { temperature: 0.0, maxOutputTokens: clamp(userOpt.maxTokens ?? 256, 16, 1024, 256) },
      cache: opts.cache,
      cacheTtlSec: opts.cacheTtlSec,
      timeoutMs: opts.timeoutMs,
      retry: opts.retry
    });
    if (!res.ok) return res.code;
    return normalizeOneCellText(res.text);
  }

  let s = input.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  s = s.trim().replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n");

  if (userOpt.removeExtraSpaces === true) s = s.replace(/\s+/g, " ").trim();

  if (userOpt.removeAccents === true) {
    try { s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); } catch { /* ignore */ }
  }

  const casing = typeof userOpt.case === "string" ? userOpt.case.toLowerCase() : "none";
  if (casing === "lower") s = s.toLowerCase();
  if (casing === "upper") s = s.toUpperCase();
  if (casing === "title") {
    s = s.toLowerCase().replace(/(^|[\s\-_/])([a-zà-öø-ÿ])/g, (m, sep, ch) => sep + ch.toUpperCase());
  }

  return normalizeOneCellText(s);
}

/** =AI.SUMMARIZE(textOrRange, [options]) */
export async function SUMMARIZE(textOrRange, options) {
  const parsed = parseOptions(options);
  if (!parsed.ok) return ERR.BAD_OPTIONS;

  const userOpt = parsed.value || {};
  const opts = normalizedCommonOptions(userOpt);

  let text = "";
  if (isMatrix(textOrRange)) text = matrixToTSV(textOrRange, LIMITS.MAX_CONTEXT_CHARS);
  else text = normalizeTextInput(textOrRange, LIMITS.MAX_INPUT_CHARS);

  text = text.trim();
  if (!text) return "";

  const maxBullets = clamp(userOpt.maxBullets, 1, 20, 6);
  const { system, user } = buildSummarizePrompt(text, { ...opts, maxBullets });

  const responseJsonSchema = {
    type: "object",
    properties: { bullets: { type: "array", items: { type: "string" } } },
    required: ["bullets"],
    additionalProperties: false
  };

  const res = await geminiGenerate({
    model: opts.model,
    system,
    user,
    generationConfig: { temperature: 0.2, maxOutputTokens: clamp(userOpt.maxTokens ?? 256, 32, 2048, 256) },
    responseMimeType: "application/json",
    responseJsonSchema,
    cache: opts.cache,
    cacheTtlSec: opts.cacheTtlSec,
    timeoutMs: opts.timeoutMs,
    retry: opts.retry
  });

  if (!res.ok) return res.code;

  const parsedOut = safeJsonParse(res.text);
  if (!parsedOut.ok || !parsedOut.value) return ERR.PARSE_ERROR;

  const bullets = Array.isArray(parsedOut.value.bullets) ? parsedOut.value.bullets : [];
  const cleaned = bullets.map((b) => normalizeOneCellText(b).replace(/\n/g, " ").trim()).filter(Boolean).slice(0, maxBullets);

  return cleaned.map((b) => `• ${b}`).join("\n");
}

/** =AI.KEY_STATUS() */
export async function KEY_STATUS() {
  const k = (await getApiKey()) || "";
  return k ? "OK" : "MISSING";
}

/** =AI.TEST([options]) */
export async function TEST(options) {
  const parsed = parseOptions(options);
  if (!parsed.ok) return ERR.BAD_OPTIONS;

  const opts = normalizedCommonOptions(parsed.value || {});
  const res = await geminiMinimalTest({ model: opts.model, timeoutMs: opts.timeoutMs });

  if (!res.ok) return res.code;
  return "OK";
}

// ID mapping (functions.json id -> implementation)
try {
  if (typeof CustomFunctions !== "undefined" && CustomFunctions.associate) {
    CustomFunctions.associate("AI.ASK", ASK);
    CustomFunctions.associate("AI.EXTRACT", EXTRACT);
    CustomFunctions.associate("AI.CLASSIFY", CLASSIFY);
    CustomFunctions.associate("AI.TRANSLATE", TRANSLATE);
    CustomFunctions.associate("AI.TABLE", TABLE);
    CustomFunctions.associate("AI.FILL", FILL);
    CustomFunctions.associate("AI.CLEAN", CLEAN);
    CustomFunctions.associate("AI.SUMMARIZE", SUMMARIZE);
    CustomFunctions.associate("AI.KEYSTATUS", KEY_STATUS);
    CustomFunctions.associate("AI.TEST", TEST);
  }
} catch (e) {
  console.error("CustomFunctions.associate failed", e);
}
