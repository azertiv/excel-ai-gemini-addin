import { DEFAULTS, ERR } from "./constants";

export function parseOptions(optionsJson) {
  if (optionsJson === undefined || optionsJson === null) return { ok: true, value: {} };

  if (typeof optionsJson !== "string") {
    try { optionsJson = String(optionsJson); }
    catch { return { ok: false, error: ERR.BAD_OPTIONS }; }
  }

  const s = optionsJson.trim();
  if (!s) return { ok: true, value: {} };

  try {
    const obj = JSON.parse(s);
    if (obj && typeof obj === "object") return { ok: true, value: obj };
    return { ok: true, value: {} };
  } catch (e) {
    return { ok: false, error: ERR.BAD_OPTIONS, message: e?.message || "Invalid JSON options" };
  }
}

export function normalizedCommonOptions(userOptions = {}) {
  const o = userOptions && typeof userOptions === "object" ? userOptions : {};
  const out = {
    lang: typeof o.lang === "string" ? o.lang : DEFAULTS.lang,
    model: typeof o.model === "string" ? o.model : undefined,
    temperature: typeof o.temperature === "number" ? o.temperature : DEFAULTS.temperature,
    maxTokens: typeof o.maxTokens === "number" ? o.maxTokens : DEFAULTS.maxTokens,
    timeoutMs: typeof o.timeoutMs === "number" ? o.timeoutMs : DEFAULTS.timeoutMs,
    retry: typeof o.retry === "number" ? Math.max(0, Math.min(3, Math.floor(o.retry))) : DEFAULTS.retry,
    cache: typeof o.cache === "string" ? o.cache : DEFAULTS.cache,
    cacheTtlSec: typeof o.cacheTtlSec === "number" ? Math.max(0, o.cacheTtlSec) : DEFAULTS.cacheTtlSec,
    debug: !!o.debug
  };

  if (typeof o.maxOutputTokens === "number") out.maxTokens = o.maxOutputTokens;
  if (typeof o.timeout === "number") out.timeoutMs = o.timeout;

  return out;
}

export function clamp(n, min, max, fallback) {
  const x = Number(n);
  if (!Number.isFinite(x)) return fallback;
  return Math.min(max, Math.max(min, x));
}
