import { ERR } from "./constants";
import { matrixToTSV, normalizeTextInput } from "./range";

export function buildAskPrompt(prompt, contextRange, options) {
  const lang = options?.lang || "fr";
  const ctx = contextRange ? matrixToTSV(contextRange, { maxChars: options?.maxContextChars }) : "";
  const system = [
    "You are an assistant embedded in Microsoft Excel custom functions.",
    `Respond in ${lang}.`,
    "Return a concise answer suitable for a single Excel cell (1 to 10 short lines).",
    "No Markdown. No code fences. No surrounding quotes.",
    "If the question cannot be answered from the provided information, say so briefly and suggest what to add."
  ].join("\n");

  const user = [ctx ? "CONTEXT (TSV):\n" + ctx : "", "USER PROMPT:\n" + normalizeTextInput(prompt)]
    .filter(Boolean)
    .join("\n\n");

  return { system, user };
}

export function parseFields(schemaOrFields) {
  const raw = (schemaOrFields ?? "").toString().trim();
  if (!raw) return { ok: false, error: ERR.BAD_SCHEMA, fields: [] };

  if (raw.startsWith("{") && raw.endsWith("}")) {
    try {
      const obj = JSON.parse(raw);
      if (obj && typeof obj === "object" && !Array.isArray(obj)) {
        const fields = Object.keys(obj).map((k) => k.trim()).filter(Boolean);
        const types = {};
        for (const k of fields) types[k] = String(obj[k]).toLowerCase();
        return { ok: true, fields, types, original: obj };
      }
    } catch { /* fallthrough */ }
  }

  const fields = raw.split(/[,;\n]+/).map((s) => s.trim()).filter(Boolean);
  return { ok: fields.length > 0, fields, types: {}, original: null, error: fields.length ? null : ERR.BAD_SCHEMA };
}

export function buildExtractPrompt(text, schemaInfo, options) {
  const lang = options?.lang || "fr";
  const fields = schemaInfo.fields || [];
  const typeHints = schemaInfo.types || {};

  const system = [
    "You extract structured fields from unstructured text.",
    "Return ONLY a JSON object that matches the provided JSON schema.",
    "Do not add extra keys.",
    "If a field is not present in the text, return an empty string for that field.",
    `Responding language for values: ${lang} (keep extracted values as-is if they are emails, phone numbers, IDs, etc.).`
  ].join("\n");

  const user = [
    "TEXT:\n" + normalizeTextInput(text),
    "FIELDS:\n" + fields.map((f) => `- ${f}${typeHints[f] ? ` (${typeHints[f]})` : ""}`).join("\n")
  ].join("\n\n");

  return { system, user };
}

export function buildTranslatePrompt(text, targetLang) {
  const system = [
    "You translate text.",
    `Translate the user content to ${targetLang}.`,
    "Preserve meaning, numbers, codes, and line breaks.",
    "Return ONLY the translated text. No quotes, no Markdown."
  ].join("\n");

  const user = normalizeTextInput(text);
  return { system, user };
}

export function buildClassifyPrompt(text, labels, options) {
  const lang = options?.lang || "fr";
  const unknownLabel = options?.unknownLabel || "UNKNOWN";
  const system = [
    "You are a strict text classifier.",
    "Choose the best label from the allowed set.",
    "Return ONLY JSON matching the schema with keys: label, confidence.",
    "confidence must be a number between 0 and 1 (higher means more certain).",
    `If you are not confident, output label "${unknownLabel}" with low confidence.`
  ].join("\n");

  const user = [
    `TEXT:\n${normalizeTextInput(text)}`,
    `LABELS:\n${labels.map((l) => `- ${l}`).join("\n")}`,
    `Output language: ${lang} (but labels must be exactly one of the allowed labels).`
  ].join("\n\n");

  return { system, user };
}

export function buildTablePrompt(prompt, contextRange, options) {
  const lang = options?.lang || "fr";
  const ctx = contextRange ? matrixToTSV(contextRange, { maxChars: options?.maxContextChars }) : "";

  const system = [
    "You generate tabular data for Excel.",
    "Return ONLY JSON matching the provided schema: {headers: string[], rows: string[][]}.",
    "headers must be an array of column names. rows must be an array of rows; each row length must equal headers length.",
    `Respond in ${lang}.`,
    "No Markdown. No additional commentary."
  ].join("\n");

  const constraints = [];
  if (options?.headers && Array.isArray(options.headers) && options.headers.length) constraints.push(`Requested headers: ${options.headers.join(" | ")}`);
  if (options?.numColumns) constraints.push(`Requested number of columns: ${options.numColumns}`);
  if (options?.maxRows) constraints.push(`Maximum rows: ${options.maxRows}`);
  if (options?.maxCols) constraints.push(`Maximum cols: ${options.maxCols}`);

  const user = [
    ctx ? "CONTEXT (TSV):\n" + ctx : "",
    "PROMPT:\n" + normalizeTextInput(prompt),
    constraints.length ? "CONSTRAINTS:\n" + constraints.map((x) => `- ${x}`).join("\n") : ""
  ].filter(Boolean).join("\n\n");

  return { system, user };
}

export function buildFillPrompt(exampleRange, targetRange, instruction, options) {
  const lang = options?.lang || "fr";
  const maxExamples = Number.isFinite(Number(options?.maxExamples)) ? Math.max(0, Math.floor(Number(options.maxExamples))) : Infinity;

  const examples = [];
  for (const row of exampleRange) {
    const r = Array.isArray(row) ? row : [row];
    const inp = (r[0] ?? "").toString().trim();
    const out = (r[1] ?? "").toString().trim();
    if (inp && out) examples.push({ input: inp, output: out });
    if (examples.length >= maxExamples) break;
  }

  const targets = [];
  for (const row of targetRange) {
    const r = Array.isArray(row) ? row : [row];
    targets.push(((r[0] ?? "")).toString().trim());
  }

  const system = [
    "You are a data transformation engine for Excel.",
    "You will learn from examples and apply the instruction to new inputs.",
    "Return ONLY JSON matching the schema: {values: string[]}.",
    "values must have the exact same number of elements as the number of target inputs, in the same order.",
    `Respond in ${lang} unless the instruction implies otherwise.`,
    "No Markdown. No extra keys.",
    "If an input is empty, return an empty string at the corresponding position."
  ].join("\n");

  const user = [
    "INSTRUCTION:\n" + normalizeTextInput(instruction),
    "EXAMPLES (input -> output):\n" + (examples.length ? examples.map((e) => `- ${e.input} -> ${e.output}`).join("\n") : "(none)"),
    "TARGET INPUTS:\n" + targets.map((t, i) => `${i + 1}. ${t || ""}`).join("\n")
  ].join("\n\n");

  return { system, user, targetCount: targets.length };
}

export function buildSummarizePrompt(text, options) {
  const lang = options?.lang || "fr";
  // No hard bullet count limits; maxOutputTokens governs length.

  const system = [
    "You summarize text for a spreadsheet cell.",
    `Respond in ${lang}.`,
    "Return a concise summary.",
    "Use bullet points with '-' when it improves readability.",
    "No Markdown. No code fences."
  ].join("\n");

  const user = "TEXT:\n" + normalizeTextInput(text);
  return { system, user };
}

export function buildCleanAiPrompt(text, options) {
  const lang = options?.lang || "fr";
  const system = [
    "You normalize text for spreadsheet usage.",
    `Respond in ${lang}.`,
    "Keep meaning. Remove redundant whitespace. Fix obvious casing issues if appropriate.",
    "Return ONLY the cleaned text. No quotes. No Markdown."
  ].join("\n");
  const user = normalizeTextInput(text);
  return { system, user };
}
