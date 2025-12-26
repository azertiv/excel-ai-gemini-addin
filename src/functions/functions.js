// src/functions/functions.js
/* global CustomFunctions */

import { geminiGenerate, geminiMinimalTest } from "../shared/gemini.js";
import { getApiKey } from "../shared/storage.js";
import { ERR, DEFAULTS, LIMITS, TOKEN_LIMITS } from "../shared/constants.js";

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
  if (typeof optionsJson === "object" && optionsJson !== null) return optionsJson;
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

function isValidHttpUrl(url) {
  if (!url || /\s/.test(url)) return false;
  try {
    const parsed = new URL(url);
    return parsed.protocol === "http:" || parsed.protocol === "https:";
  } catch {
    return false;
  }
}

function errorCode(code) {
  return code || ERR.API_ERROR;
}

function normalizeNewlines(s) {
  return String(s || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

function extractFormula(text) {
  const cleaned = normalizeNewlines(String(text || "").replace(/```[a-z]*\n?/gi, "").replace(/```/g, ""));
  if (!cleaned) return "";

  const lines = cleaned
    .split(/\n+/)
    .map((l) => l.trim())
    .filter(Boolean);

  for (const line of lines) {
    const sanitized = line.replace(/^[\s>*•-]+/, "").replace(/^['\"]/, "");
    if (sanitized.startsWith("=")) return sanitized;
    const match = sanitized.match(/=(.+)/);
    if (match) return "=" + match[1].trim();
  }

  const first = lines[0] || "";
  return first ? (first.startsWith("=") ? first : "=" + first) : "";
}

function matrixToTSV(matrix, maxChars) {
  if (!Array.isArray(matrix)) return "";

  // No hard cap by default. If maxChars is provided, it is treated as an optional safety limit.
  const limit = Number.isFinite(Number(maxChars)) && Number(maxChars) > 0 ? Math.floor(Number(maxChars)) : Infinity;

  let out = "";
  for (const row of matrix) {
    if (out.length >= limit) break;
    const r = Array.isArray(row) ? row : [row];
    const line = r.map((c) => safeString(c).replace(/\t/g, " ").replace(/\n/g, " ")).join("\t");
    const next = out ? out + "\n" + line : line;
    if (next.length > limit) {
      out = next.slice(0, limit);
      break;
    }
    out = next;
  }
  return out;
}

function flattenLabels(labelsOrRange) {
  if (labelsOrRange === null || labelsOrRange === undefined) return [];
  if (typeof labelsOrRange === "string") {
    const s = labelsOrRange.trim();
    if (!s) return [];
    const parts = s.includes("|") ? s.split("|") : s.split(/[,;\n]+/);
    return parts.map((x) => x.trim()).filter(Boolean);
  }
  if (Array.isArray(labelsOrRange)) {
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
    if (v.length === 0) return "";
    if (Array.isArray(v[0])) return matrixToTSV(v);
    return v.map((x) => safeString(x)).join("\n");
  }
  return safeString(v);
}

function normalizeRangeToMatrix(v) {
  if (!Array.isArray(v)) return [[v]];
  if (Array.isArray(v[0])) return v.map((row) => (Array.isArray(row) ? row : [row]));
  return v.map((cell) => [cell]);
}

function lightlyCleanExtractedValue(value, instruction) {
  const raw = safeString(value).trim();
  if (!raw) return "";

  const loweredInstruction = safeString(instruction).toLowerCase();
  const isEmail = loweredInstruction.includes("mail") || loweredInstruction.includes("email");

  if (isEmail) {
    let s = raw;
    s = s.replace(/\s*\[\s*at\s*\]\s*|\s*\(\s*at\s*\)\s*|\s+at\s+/gi, "@");
    s = s.replace(/\s*\[\s*dot\s*\]\s*|\s*\(\s*dot\s*\)\s*|\s+dot\s+/gi, ".");
    s = s.replace(/\s*\[\s*point\s*\]\s*|\s*\(\s*point\s*\)\s*|\s+point\s+/gi, ".");
    s = s.replace(/\s*@\s*/g, "@");
    s = s.replace(/\s*\.\s*/g, ".");
    s = s.replace(/[<>\"'`\u201c\u201d]/g, "");
    s = s.replace(/\s+/g, " ").trim().toLowerCase();
    return s;
  }

  return raw;
}

function isEmailLikeInstruction(instruction) {
  const s = safeString(instruction).toLowerCase();
  // Common French/English phrasings.
  return (
    s.includes("email") ||
    s.includes("e-mail") ||
    s.includes("mail") ||
    s.includes("courriel") ||
    s.includes("adresse électronique") ||
    s.includes("adresse electronique") ||
    s.includes("adresse email") ||
    s.includes("adresse mail")
  );
}

function extractEmailsFromText(text, instructionHint = "email") {
  const cleaned = lightlyCleanExtractedValue(text, instructionHint);
  // RFC-like pragmatic regex (good enough for spreadsheet extraction).
  const re = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  const matches = cleaned.match(re) || [];
  // Deduplicate while preserving order.
  const seen = new Set();
  const out = [];
  for (const m of matches) {
    const v = safeString(m).trim();
    if (!v) continue;
    const key = v.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(v);
  }
  return out;
}

function truncateForCell(s, maxChars = LIMITS.MAX_CELL_CHARS) {
  const t = normalizeNewlines(s);
  if (t.length <= maxChars) return t;
  return t.slice(0, maxChars) + "\n…(truncated)";
}

function normalizeMatrixInput(value) {
  if (Array.isArray(value)) return normalizeRangeToMatrix(value);
  return [[value]];
}

function fillMatrix(matrix, value) {
  if (!Array.isArray(matrix)) return [[value]];
  return matrix.map((row) => {
    if (Array.isArray(row)) return row.map(() => value);
    return [value];
  });
}

function extractJsonObject(text) {
  const s = String(text || "").trim();
  if (!s) return null;

  // Nettoyage Markdown ```json ... ```
  const clean = s.replace(/```json/g, "").replace(/```/g, "").trim();

  // Try parsing cleaned string
  const p1 = safeJsonParse(clean);
  if (p1.ok && p1.value && typeof p1.value === "object") return p1.value;

  // Fallback: Try to find brackets
  const start = clean.indexOf("{");
  const end = clean.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const candidate = clean.slice(start, end + 1);
    const p2 = safeJsonParse(candidate);
    if (p2.ok && p2.value && typeof p2.value === "object") return p2.value;
  }
  return null;
}

// ---------- per-function prompting ----------

function sysAsk(lang = "fr") {
  return [
    "You are an assistant embedded in Microsoft Excel custom functions.",
    `Respond in ${lang}.`,
    "Return a clear and accurate answer suitable for an Excel cell.",
    // Suppression de "concise 1-10 lines" pour permettre des réponses longues si tokens > 256
    "No Markdown. No code fences. No surrounding quotes.",
    "If the question cannot be answered from the provided information, say so briefly."
  ].join("\n");
}

function sysTranslate(targetLang, expectedItems) {
  const n = Number.isFinite(Number(expectedItems)) ? Math.floor(Number(expectedItems)) : 0;
  if (n > 1) {
    return [
      "You are a translation engine.",
      `Translate each cell independently into ${targetLang}.`,
      "Return STRICT JSON only (no Markdown, no code fences).",
      `Return an object with a single key 'items' containing exactly ${n} strings in the same order as the provided cells.`,
      "Each item must contain ONLY the translated text for that cell.",
      "Preserve numbers, units, and identifiers as-is unless they require translation.",
      "For empty inputs, return an empty string at the same position.",
      "Do not add any keys other than 'items'."
    ].join("\n");
  }

  return [
    "You are a translation engine.",
    `Translate the user text into ${targetLang}.`,
    "Return only the translated text. No quotes. No explanations."
  ].join("\n");
}

function sysClassify(labels, lang = "en") {
  return [
    "You are a strict classifier.",
    `Return exactly one label from this set, using the label text verbatim: ${labels.join(" | ")}`,
    "Do not translate, expand, or paraphrase labels.",
    "If uncertain, return exactly: UNKNOWN",
    `Respond in ${lang}.`,
    "Return only the label.",
  ].join("\n");
}
function sysClean(lang = "fr", expectedItems) {
  if (typeof expectedItems === "number" && expectedItems > 1) {
    return [
      "You are a text normalizer for spreadsheet cells.",
      `Respond in ${lang}.`,
      "Return STRICT JSON only (no Markdown, no code fences).",
      `Return an object with a single key 'items' containing exactly ${expectedItems} strings in the same order as the provided cells.`,
      "Preserve the meaning of each cell independently.",
      "For empty inputs, return an empty string at the same position.",
      "Do not invent or merge content."
    ].join("\n");
  }

  return [
    "You are a text normalizer for spreadsheet cells.",
    `Respond in ${lang}.`,
    "Return only the cleaned text.",
    "No quotes. No explanations."
  ].join("\n");
}

function sysConsistent(lang = "fr", expectedItems) {
  return [
    "You harmonize spreadsheet entries that refer to the same real-world value.",
    `Respond in ${lang}.`,
    "Return STRICT JSON only (no Markdown, no code fences).",
    `Return an object with a single key 'items' containing exactly ${expectedItems} strings in the same order as the provided cells.`,
    "Normalize casing, accents, spacing, and fix obvious typos.",
    "When several cells refer to the same entity, use ONE consistent, best-written value for all of them.",
    "Keep outputs aligned with inputs; do not merge or reorder rows.",
    "If an input is empty or whitespace-only, return an empty string for that position.",
    "Do not invent new information beyond correcting the given values."
  ].join("\n");
}

function sysSummarize(lang = "fr") {
  return [
    "You summarize text for a spreadsheet cell.",
    `Respond in ${lang}.`,
    "Return a clear summary suitable for an Excel cell.",
    "Use bullet points with '-' when it improves readability.",
    "No Markdown headers. No code fences. No surrounding quotes."
  ].join("\n");
}

function sysExtract(instruction, lang = "fr", expectedItems) {
  const strictArray =
    typeof expectedItems === "number" && expectedItems > 0
      ? `Return an object with a single key 'items' which is an array of exactly ${expectedItems} strings, preserving order.`
      : "Return an object with a single key 'items' which is an array of strings.";

  return [
    "You are an expert extraction engine.",
    `Goal: Extract all entities matching this description: "${instruction}"`,
    `Respond in ${lang}.`,
    "Lightly normalize results (trim spaces, fix obvious email obfuscation like [at]/(at) -> @ and [dot]/(dot)/point -> .).",
    "Return STRICT JSON only (no Markdown, no code fences).",
    strictArray,
    "Example: { \"items\": [\"match1\", \"match2\"] }",
    typeof expectedItems === "number"
      ? "If a value is missing for a cell, return an empty string in that position."
      : "If nothing found, return { \"items\": [] }.",
    "Extract exact values from the text without inventing data."
  ].join("\n");
}

function sysFill(lang = "fr", expectedItems) {
  const n = Number.isFinite(Number(expectedItems)) ? Math.floor(Number(expectedItems)) : 0;
  const spec = n > 0
    ? `Return an object with a single key 'values' containing exactly ${n} strings, in the same order as the target rows.`
    : "Return an object with a single key 'values' containing an array of strings.";

  return [
    "You are filling spreadsheet cells based on examples.",
    `Respond in ${lang}.`,
    "Return STRICT JSON only (no Markdown, no code fences).",
    spec,
    "Return only the values, in order, one per target row.",
    "If a value is unknown, return an empty string for that row.",
    "Do not add any keys other than 'values'."
  ].join("\n");
}

function sysFormula(lang) {
  const isFr = (lang || "").toLowerCase().startsWith("fr");
  return [
    "You are an expert Excel formula generator.",
    "Your goal is to output a VALID Excel formula string based on the user request.",
    "Leverage advanced Excel capabilities (dynamic arrays, LET/LAMBDA, structured references, advanced date/time, lookup, statistics, financial functions) when relevant.",
    `Respond in ${lang}.`,
    isFr
      ? "Use FRENCH Excel function names (e.g., SOMME, SI, RECHERCHEV...)."
      : "Use ENGLISH Excel function names (e.g., SUM, IF, VLOOKUP...).",
    isFr ? "Use SEMICOLON (;) as argument separator." : "Use COMMA (,) as argument separator.",
    "Return exactly one ready-to-use Excel formula with no surrounding text.",
    "Return ONLY the formula starting with '='.",
    "No Markdown. No code fences. No explanations."
  ].join("\n");
}

  function sysWeb(lang = "fr") {
    return [
      "You are a meticulous fact-finding assistant with access to reliable web knowledge.",
      "Return only one precise, up-to-date factual value plus the best authoritative source URL.",
      "Never fabricate numbers or URLs. Use official or authoritative sources only.",
      "Match the requested timeframe and scope exactly; ignore partial or approximate figures.",
      "If the data cannot be confirmed with high confidence, return empty strings and explain why in a 'reason' field.",
      `Respond in ${lang}.`,
      "Return STRICT JSON only (no Markdown, no code fences).",
      'Schema: {"value": "<concise value>", "source": "https://...", "reason": "<why unavailable>"}.',
      "The value must mirror the source exactly and stay under 80 characters."
    ].join("\n");
  }

  // ---------- core call wrapper ----------

async function callGemini({ system, user, options, functionName }) {
  const opt = options || {};
  const temperature = typeof opt.temperature === "number"
    ? clamp(opt.temperature, 0, 1, DEFAULTS.temperature)
    : DEFAULTS.temperature;

  // maxOutputTokens handling:
  // - If the user provides an explicit per-formula value, clamp it to the UI bounds.
  // - Otherwise do NOT set maxOutputTokens here, so the stored global setting (taskpane slider) is used.
  const maxTokensCandidate = opt.maxTokens ?? opt.maxOutputTokens;
  let maxOutputTokens;
  if (maxTokensCandidate !== undefined) {
    const n = Math.floor(Number(maxTokensCandidate));
    if (Number.isFinite(n)) {
      maxOutputTokens = Math.min(TOKEN_LIMITS.MAX, Math.max(TOKEN_LIMITS.MIN, n));
    }
  }

  const timeoutMs = typeof opt.timeoutMs === "number"
    ? clamp(opt.timeoutMs, 1000, 120000, DEFAULTS.timeoutMs)
    : DEFAULTS.timeoutMs;

  const retry = typeof opt.retry === "number"
    ? clamp(opt.retry, 0, 2, DEFAULTS.retry)
    : DEFAULTS.retry;

  const cacheMode = typeof opt.cache === "string" ? opt.cache : DEFAULTS.cache;
  const cacheTtlSec = typeof opt.cacheTtlSec === "number" ? clamp(opt.cacheTtlSec, 0, 24 * 3600, DEFAULTS.cacheTtlSec) : DEFAULTS.cacheTtlSec;
  const cacheOnly = Boolean(opt.cacheOnly);

  const generationConfig = { temperature };
  if (typeof maxOutputTokens === "number") generationConfig.maxOutputTokens = maxOutputTokens;

  const res = await geminiGenerate({
    model: opt.model,
    system,
    user,
    generationConfig,
    tools: Array.isArray(opt.tools) ? opt.tools : undefined,
    cache: cacheMode,
    cacheTtlSec,
    cacheOnly,
    timeoutMs,
    retry,
    responseMimeType: opt.responseMimeType,
    responseJsonSchema: opt.responseJsonSchema,
    functionName: functionName
  });

  return res;
}

// ---------- Custom Functions ----------

export async function KEY_STATUS() {
  try {
    const key = await getApiKey();
    return key ? "OK" : "MISSING";
  } catch {
    return "MISSING";
  }
}

export async function TEST() {
  try {
    const res = await geminiMinimalTest({ timeoutMs: DEFAULTS?.timeoutMs || 15000 });
    if (!res.ok) return errorCode(res.code);
    return "OK";
  } catch (e) {
    return fillMatrix(normalizeRangeToMatrix(text), errorCode(ERR.API_ERROR));
  }
}

export async function ASK(prompt, contextRange, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";

    const ctx = contextRange ? matrixToTSV(contextRange, opt.maxContextChars) : "";
    const user = [ctx ? `CONTEXT (TSV):\n${ctx}` : "", `USER PROMPT:\n${coerceToTextOrJoin2D(prompt)}`]
      .filter(Boolean)
      .join("\n\n");

    const res = await callGemini({
      system: sysAsk(lang),
      user,
      options: opt,
      functionName: "AI.ASK"
    });

    if (!res.ok) return errorCode(res.code);
    return truncateForCell(res.text);
  } catch (e) {
    return fillMatrix(normalizeRangeToMatrix(text), errorCode(ERR.API_ERROR));
  }
}

export async function WEB(prompt, focusRange, showSource) {
  try {
    const query = normalizeNewlines(coerceToTextOrJoin2D(prompt)).trim();
    if (!query) return errorCode(ERR.BAD_INPUT);

    const focus = focusRange ? normalizeNewlines(coerceToTextOrJoin2D(focusRange)).trim() : "";
    
    // Détection booléenne simple pour l'option lien
    const wantsHyperlink = (() => {
      if (!showSource) return false;
      const s = String(showSource).toLowerCase().trim();
      return s === "1" || s === "true" || s === "yes" || s === "oui";
    })();

    const user = [
      `QUESTION: ${query}`,
      focus ? `FOCUS / ENTITY: ${focus}` : "",
      "Return STRICT JSON with the confirmed value and the best source URL.",
      "If the value cannot be confirmed, leave value empty.",
      'Schema: {"value": "...", "source": "...", "reason": "..."}'
    ].join("\n\n");

    const res = await callGemini({
      system: sysWeb("fr"),
      user,
      options: {
        temperature: 0.0,
        responseMimeType: "application/json",
        // [CORRECTION] Syntaxe standard pour l'API publique v1beta
        tools: [{ googleSearch: {} }] 
      },
      functionName: "AI.WEB"
    });

    if (!res.ok) return errorCode(res.code);

    const obj = extractJsonObject(res.text);
    if (!obj) return errorCode(ERR.PARSE_ERROR);

    let value = safeString(obj.value).trim();
    let source = safeString(obj.source).trim();
    const reason = safeString(obj.reason).trim();

    // [CORRECTION] Si l'IA n'a pas rempli la source dans le JSON, on la prend dans les métadonnées techniques
    if (!isValidHttpUrl(source) && res.groundingMetadata?.groundingChunks?.length > 0) {
        // On prend la première source Web disponible
        const firstWebChunk = res.groundingMetadata.groundingChunks.find(c => c.web?.uri);
        if (firstWebChunk) {
            source = firstWebChunk.web.uri;
        }
    }

    if (!value) {
      return reason || errorCode(ERR.NOT_FOUND);
    }

    value = truncateForCell(value);

    // Si l'utilisateur veut un lien et qu'on a une URL valide
    if (wantsHyperlink && isValidHttpUrl(source)) {
      // Échappement des guillemets pour la formule Excel
      const escVal = value.replace(/"/g, '""');
      const escUrl = source.replace(/"/g, '""');
      return `=LIEN_HYPERTEXTE("${escUrl}";"${escVal}")`;
    }

    return value;
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

export async function TRANSLATE(text, targetLang, options) {
  try {
    const opt = parseOptions(options);
    const lang = safeString(targetLang).trim() || "en";
    const matrix = normalizeRangeToMatrix(text);
    const flatCells = [];
    for (const row of matrix) {
      for (const cell of row) {
        flatCells.push(normalizeNewlines(coerceToTextOrJoin2D(cell)));
      }
    }

    if (flatCells.length === 0) return fillMatrix(matrix, errorCode(ERR.BAD_INPUT));

    // Single cell => single translation (still returned as a 1x1 matrix so it can also spill when used on ranges).
    if (flatCells.length === 1) {
      const raw = flatCells[0];
      if (!raw.trim()) return [[""]];

      const res = await callGemini({
        system: sysTranslate(lang, 1),
        user: raw,
        options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.2 },
        functionName: "AI.TRANSLATE"
      });

      if (!res.ok) return fillMatrix(matrix, errorCode(res.code));
      return [[truncateForCell(res.text)]];
    }

    const hasContent = flatCells.some((cell) => safeString(cell).trim());
    if (!hasContent) return matrix.map((row) => row.map(() => ""));

    const userCells = flatCells
      .map((cell, idx) => `${idx + 1}. ${cell ? cell : "<empty>"}`)
      .join("\n");

    const res = await callGemini({
      system: sysTranslate(lang, flatCells.length),
      user: ["Cells:", userCells].join("\n"),
      options: {
        ...opt,
        temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0,
        responseMimeType: "application/json"
      },
      functionName: "AI.TRANSLATE"
    });

    if (!res.ok) return fillMatrix(matrix, errorCode(res.code));

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.items)) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));
    if (obj.items.length !== flatCells.length) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));

    let idx = 0;
    return matrix.map((row) => row.map(() => truncateForCell(safeString(obj.items[idx++]))));
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

export async function CLASSIFY(text, labels, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "en";
    const threshold = typeof opt.threshold === "number" ? clamp(opt.threshold, 0, 1, 0.55) : 0.55;

    const labs = Array.from(
      new Map(
        flattenLabels(labels)
          .map((label) => safeString(label).trim())
          .filter(Boolean)
          .map((label) => [label.toLowerCase(), label])
      ).values()
    );
    const matrix = normalizeRangeToMatrix(text);
    const flatCells = [];
    for (const row of matrix) {
      for (const cell of row) {
        flatCells.push(normalizeNewlines(coerceToTextOrJoin2D(cell)));
      }
    }

    if (!labs.length) return fillMatrix(matrix, errorCode(ERR.BAD_INPUT));
    if (flatCells.length === 0) return fillMatrix(matrix, errorCode(ERR.BAD_INPUT));

    const normalizeLabel = (value) => {
      const raw = safeString(value).trim();
      if (!raw) return "UNKNOWN";
      if (raw.toUpperCase() === "UNKNOWN") return "UNKNOWN";
      const match = labs.find((l) => l.toLowerCase() === raw.toLowerCase());
      return match || "UNKNOWN";
    };

    if (flatCells.length === 1) {
      const raw = flatCells[0];
      const system = sysClassify(labs, lang);
      const user = [
        "TEXT:",
        raw,
        "",
        `Return only one label. If confidence < ${threshold}, return UNKNOWN.`
      ].join("\n");

      const res = await callGemini({ system, user, options: opt, functionName: "AI.CLASSIFY" });
      if (!res.ok) return errorCode(res.code);

      return [[normalizeLabel(res.text)]];
    }

    const hasContent = flatCells.some((cell) => safeString(cell).trim());
    if (!hasContent) return matrix.map((row) => row.map(() => "UNKNOWN"));

    const userCells = flatCells
      .map((cell, idx) => `${idx + 1}. ${cell ? cell : "<empty>"}`)
      .join("\n");

    const system = [
      "You are a strict classifier.",
      `You will classify ${flatCells.length} independent cell values.`,
      `Labels: ${labs.join(" | ")}`,
      "Use the label text verbatim; do not translate or paraphrase labels.",
      `If confidence < ${threshold} or information is missing, return exactly: UNKNOWN`,
      `Respond in ${lang}.`,
      "Return STRICT JSON only (no Markdown, no code fences).",
      `Return an object with a single key 'items' containing exactly ${flatCells.length} strings in the same order as the provided cells.`,
      "Each item must be one of the provided labels or UNKNOWN.",
      "No explanations.",
    ].join("\n");

    const user = ["Cells:", userCells].join("\n");

    const res = await callGemini({
      system,
      user,
      options: {
        ...opt,
        temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0,
        responseMimeType: "application/json"
      },
      functionName: "AI.CLASSIFY"
    });

    if (!res.ok) return fillMatrix(matrix, errorCode(res.code));

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.items)) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));
    if (obj.items.length !== flatCells.length) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));

    let idx = 0;
    return matrix.map((row) => row.map(() => normalizeLabel(obj.items[idx++])));
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

export async function CLEAN(text, options) {
  try {
    const opt = parseOptions(options);

    const matrix = normalizeRangeToMatrix(text);
    const flatCells = [];
    for (const row of matrix) {
      for (const cell of row) {
        flatCells.push(normalizeNewlines(coerceToTextOrJoin2D(cell)));
      }
    }

    if (flatCells.length === 0) return fillMatrix(matrix, errorCode(ERR.BAD_INPUT));

    if (flatCells.length === 1) {
      const raw = flatCells[0];
      if (!raw.trim()) return [[""]];

      const lang = opt.lang || "fr";
      const res = await callGemini({
        system: sysClean(lang),
        user: raw,
        options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0 },
        functionName: "AI.CLEAN"
      });

      if (!res.ok) return fillMatrix(matrix, errorCode(res.code));
      return [[truncateForCell(res.text)]];
    }

    const lang = opt.lang || "fr";
    const hasContent = flatCells.some((cell) => safeString(cell).trim());
    if (!hasContent) return matrix.map((row) => row.map(() => ""));

    const userCells = flatCells
      .map((cell, idx) => `${idx + 1}. ${cell ? cell : "<empty>"}`)
      .join("\n");

    const user = [
      `You will clean ${flatCells.length} independent cell values.`,
      "Return STRICT JSON only (no Markdown, no code fences).",
      `Return an object with a single key 'items' containing exactly ${flatCells.length} strings in the same order as the provided cells.`,
      "Preserve the intent of each cell; do not merge or summarize.",
      "Use an empty string for empty or whitespace-only inputs.",
      "Lightly normalize whitespace and punctuation without inventing content.",
      "Cells:",
      userCells
    ].join("\n");

    const res = await callGemini({
      system: sysClean(lang, flatCells.length),
      user,
      options: {
        ...opt,
        temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0,
        responseMimeType: "application/json"
      },
      functionName: "AI.CLEAN"
    });

    if (!res.ok) return fillMatrix(matrix, errorCode(res.code));

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.items)) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));
    if (obj.items.length !== flatCells.length) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));

    const cleaned = obj.items.map((item) => safeString(item));
    let idx = 0;
    return matrix.map((row) => row.map(() => truncateForCell(cleaned[idx++])));
  } catch (e) {
    return fillMatrix(normalizeRangeToMatrix(text), errorCode(ERR.API_ERROR));
  }
}

export async function CONSISTENT(text, options) {
  try {
    const opt = parseOptions(options);

    const matrix = normalizeRangeToMatrix(text);
    const flatCells = [];
    for (const row of matrix) {
      for (const cell of row) {
        flatCells.push(normalizeNewlines(coerceToTextOrJoin2D(cell)));
      }
    }

    if (flatCells.length === 0) return fillMatrix(matrix, errorCode(ERR.BAD_INPUT));

    const hasContent = flatCells.some((cell) => safeString(cell).trim());
    if (!hasContent) return matrix.map((row) => row.map(() => ""));

    const userCells = flatCells
      .map((cell, idx) => `${idx + 1}. ${cell ? cell : "<vide>"}`)
      .join("\n");

    const lang = opt.lang || "fr";
    const user = [
      `You will harmonize ${flatCells.length} cell values for consistent sorting/counting.`,
      "Cells:",
      userCells
    ].join("\n");

    const res = await callGemini({
      system: sysConsistent(lang, flatCells.length),
      user,
      options: {
        ...opt,
        temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0,
        responseMimeType: "application/json"
      },
      functionName: "AI.CONSISTENT"
    });

    if (!res.ok) return fillMatrix(matrix, errorCode(res.code));

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.items)) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));
    if (obj.items.length !== flatCells.length) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));

    const normalized = obj.items.map((item) => safeString(item));
    let idx = 0;
    return matrix.map((row) => row.map(() => truncateForCell(normalized[idx++])));
  } catch (e) {
    return fillMatrix(normalizeRangeToMatrix(text), errorCode(ERR.API_ERROR));
  }
}

export async function SUMMARIZE(textOrRange, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";
    const raw = normalizeNewlines(coerceToTextOrJoin2D(textOrRange));
    if (!raw.trim()) return "";

    const res = await callGemini({
      system: sysSummarize(lang),
      user: raw,
      options: { ...opt, temperature: typeof opt.temperature === "number" ? opt.temperature : 0.2 },
      functionName: "AI.SUMMARIZE"
    });

    if (!res.ok) return errorCode(res.code);
    return truncateForCell(res.text);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

export async function EXTRACT(textOrRange, instruction, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";

    const instr = safeString(instruction).trim();
    if (!instr) return errorCode(ERR.BAD_INPUT);

    const matrix = normalizeRangeToMatrix(textOrRange);
    const flatCells = [];
    for (const row of matrix) {
      for (const cell of row) {
        flatCells.push(normalizeNewlines(coerceToTextOrJoin2D(cell)));
      }
    }

    if (flatCells.length === 0) return fillMatrix(matrix, errorCode(ERR.BAD_INPUT));

    // Fast-path: email extraction is far more reliable (and cheaper) with a deterministic regex.
    if (isEmailLikeInstruction(instr)) {
      if (flatCells.length === 1) {
        const emails = extractEmailsFromText(flatCells[0], instr);
        if (!emails.length) return [[errorCode(ERR.NOT_FOUND)]];
        return emails.map((e) => [truncateForCell(e)]);
      }

      let i = 0;
      return matrix.map((row) =>
        row.map(() => {
          const emails = extractEmailsFromText(flatCells[i++], instr);
          if (!emails.length) return errorCode(ERR.NOT_FOUND);
          return truncateForCell(emails[0]);
        })
      );
    }

    if (flatCells.length === 1) {
      const raw = flatCells[0];
      if (!raw.trim()) return [[errorCode(ERR.NOT_FOUND)]];

      const res = await callGemini({
        system: sysExtract(instr, lang),
        user: raw,
        options: {
          ...opt,
          temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0,
          responseMimeType: "application/json"
        },
        functionName: "AI.EXTRACT"
      });

      if (!res.ok) return [[errorCode(res.code)]];

      const obj = extractJsonObject(res.text);
      if (!obj || !Array.isArray(obj.items)) return [[errorCode(ERR.PARSE_ERROR)]];

      const items = obj.items
        .map((item) => lightlyCleanExtractedValue(item, instr))
        .map((x) => safeString(x))
        .filter((x) => x.trim());

      if (items.length === 0) return [[errorCode(ERR.NOT_FOUND)]];
      return items.map((v) => [truncateForCell(v)]);
    }

    const hasNonEmptyCell = flatCells.some((cell) => safeString(cell).trim());
    if (!hasNonEmptyCell) return matrix.map((row) => row.map(() => errorCode(ERR.NOT_FOUND)));

    const userCells = flatCells
      .map((cell, idx) => `${idx + 1}. ${cell ? cell : "<empty>"}`)
      .join("\n");

    const user = [
      `You will process ${flatCells.length} independent cell values.`,
      `Instruction: "${instr}".`,
      "Return STRICT JSON only (no Markdown, no code fences).",
      `Return an object with a single key 'items' containing exactly ${flatCells.length} strings in the same order as the provided cells.`,
      "Use an empty string when the requested value is absent or uncertain for a cell.",
      "Do not invent values; only return data present in the corresponding cell.",
      "Lightly clean outputs (trim spaces, fix obvious email obfuscation).",
      "Cells:",
      userCells
    ].join("\n");

    const res = await callGemini({
      system: sysExtract(instr, lang, flatCells.length),
      user,
      options: {
        ...opt,
        temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0,
        responseMimeType: "application/json"
      },
      functionName: "AI.EXTRACT"
    });

    if (!res.ok) return fillMatrix(matrix, errorCode(res.code));

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.items)) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));
    if (obj.items.length !== flatCells.length) return fillMatrix(matrix, errorCode(ERR.PARSE_ERROR));

    const cleaned = obj.items.map((item) => lightlyCleanExtractedValue(item, instr));
    let idx = 0;
    return matrix.map((row) =>
      row.map(() => {
        const v = cleaned[idx++];
        if (safeString(v).trim()) return truncateForCell(v);
        return errorCode(ERR.NOT_FOUND);
      })
    );
  } catch (e) {
    return fillMatrix(normalizeRangeToMatrix(textOrRange), errorCode(ERR.API_ERROR));
  }
}

/**
 * =AI.TABLE(prompt, [contextRange], [options])
 * Correctif : Utilise responseMimeType: application/json pour éviter les erreurs de format,
 * et normalise la matrice (Spill) pour qu'elle soit rectangulaire.
 */
export async function TABLE(prompt, contextRange, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";
    const maxRows = (() => {
      const n = Math.floor(Number(opt.maxRows));
      return Number.isFinite(n) && n > 0 ? n : Infinity;
    })();
    // On retire Headers de opt ici car on veut que l'IA les génère si non fournis,
    // mais on les passe au prompt system.
    const requestedHeaders = Array.isArray(opt.headers) ? opt.headers.map((h) => safeString(h)) : null;

    const ctx = contextRange ? matrixToTSV(contextRange, opt.maxContextChars) : "";

    // System Prompt forcé en mode JSON strict
    const system = [
      "You are a generator of tabular data for Excel.",
      `Respond in ${lang}.`,
      "Return STRICT JSON only.",
      "The JSON must have this schema: { \"headers\": [string], \"rows\": [[string]] }.",
      "rows must be an array of arrays. Each inner array must have the exact same length as headers.",
      requestedHeaders ? `Use these headers exactly: ${requestedHeaders.join(", ")}` : "",
      "No Markdown. No code fences."
    ].filter(Boolean).join("\n");

    const user = [
      ctx ? `CONTEXT (TSV):\n${ctx}` : "",
      `PROMPT:\n${coerceToTextOrJoin2D(prompt)}`
    ].filter(Boolean).join("\n\n");

    // APPEL avec responseMimeType 'application/json'
    const res = await callGemini({
      system,
      user,
      options: {
        ...opt,
        temperature: typeof opt.temperature === "number" ? opt.temperature : 0.1,
        responseMimeType: "application/json"
      },
      functionName: "AI.TABLE"
    });

    if (!res.ok) return errorCode(res.code);

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.headers) || !Array.isArray(obj.rows)) {
      return errorCode(ERR.PARSE_ERROR);
    }

    const h = obj.headers.map((x) => safeString(x));
    if (h.length === 0) return errorCode(ERR.PARSE_ERROR);

    // Construction de la matrice de sortie (Rectangulaire)
    const out = [h];
    const numCols = h.length;

    const rows = obj.rows.slice(0, maxRows);
    for (const r of rows) {
      if (!Array.isArray(r)) continue;
      // Normalisation: on s'assure que la ligne a exactement numCols colonnes
      const cleanRow = [];
      for(let i=0; i<numCols; i++) {
        const cell = r[i] === null || r[i] === undefined ? "" : safeString(r[i]);
        cleanRow.push(truncateForCell(cell));
      }
      out.push(cleanRow);
    }

    return out;
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

export async function FILL(exampleRange, targetRange, instruction, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";

    const examples = Array.isArray(exampleRange) ? exampleRange : [];
    const targets = Array.isArray(targetRange) ? targetRange : [];

    const targetRows = Array.isArray(targets) ? targets.length : 0;
    if (!targetRows) return [];

    // No hard limit: fill the entire target range by default.
    // If the user provides opt.maxRows, respect it (still only bounded by tokens/model limits).
    const rowsToFill = (() => {
      const n = Math.floor(Number(opt.maxRows));
      if (Number.isFinite(n) && n > 0) return Math.min(targetRows, n);
      return targetRows;
    })();

    const exTSV = matrixToTSV(examples, opt.maxExamplesChars);
    const tgtTSV = matrixToTSV(targets.slice(0, rowsToFill), opt.maxTargetsChars);

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
      options: {
        ...opt,
        temperature: typeof opt.temperature === "number" ? opt.temperature : 0.0,
        responseMimeType: "application/json"
      },
      functionName: "AI.FILL"
    });

    if (!res.ok) return errorCode(res.code);

    const obj = extractJsonObject(res.text);
    if (!obj || !Array.isArray(obj.values)) return errorCode(ERR.PARSE_ERROR);

    const values = obj.values.map((x) => safeString(x));
    while (values.length < rowsToFill) values.push("");
    return values.slice(0, rowsToFill).map((v) => [truncateForCell(v)]);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

export async function FORMULA(instruction, contextRange, options) {
  try {
    const opt = parseOptions(options);
    const lang = opt.lang || "fr";

    const ctx = contextRange ? matrixToTSV(contextRange, opt.maxContextChars) : "";
    const user = [
      ctx ? `CONTEXT (TSV):\n${ctx}` : "",
      `INSTRUCTION:\n${coerceToTextOrJoin2D(instruction)}`
    ]
      .filter(Boolean)
      .join("\n\n");

    const res = await callGemini({
      system: sysFormula(lang),
      user,
      options: {
        ...opt,
        temperature: 0.0 // Strict as requested
      },
      functionName: "AI.FORMULA"
    });

    if (!res.ok) return errorCode(res.code);

    const formula = extractFormula(res.text);
    if (!formula) return errorCode(ERR.PARSE_ERROR);

    return truncateForCell(formula);
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

export function COUNT(range, valueToCount) {
  try {
    const matrix = normalizeMatrixInput(range);
    const target = safeString(valueToCount).trim();
    if (!target) return 0;

    const targetLower = target.toLowerCase();
    let total = 0;

    for (const row of matrix) {
      if (!Array.isArray(row)) continue;
      for (const cell of row) {
        const candidate = safeString(cell).trim();
        if (!candidate) continue;
        if (candidate.toLowerCase() === targetLower) total += 1;
      }
    }

    return total;
  } catch (e) {
    return errorCode(ERR.API_ERROR);
  }
}

function registerCustomFunctions() {
  if (typeof CustomFunctions === "undefined" || typeof CustomFunctions.associate !== "function") return false;

  const pairs = [
    ["AI.ASK", ASK],
    ["AI.WEB", WEB],
    ["AI.EXTRACT", EXTRACT],
    ["AI.CLASSIFY", CLASSIFY],
    ["AI.TRANSLATE", TRANSLATE],
    ["AI.TABLE", TABLE],
    ["AI.FILL", FILL],
    ["AI.FORMULA", FORMULA],
    ["AI.COUNT", COUNT],
    ["AI.CONSISTENT", CONSISTENT],
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
      try { console.warn(`[AI] CustomFunctions.associate failed for ${id}`, e); } catch { }
    }
  }

  return any;
}

const _registered = registerCustomFunctions();

if (!_registered && typeof setTimeout === "function") {
  let attempts = 0;
  const maxAttempts = 20;
  const intervalMs = 500;

  const tick = () => {
    attempts++;
    if (registerCustomFunctions() || attempts >= maxAttempts) return;
    setTimeout(tick, intervalMs);
  };

  setTimeout(tick, 0);
}
