// src/shared/openai.js

import { OPENAI, DEFAULTS, LIMITS, TOKEN_LIMITS, ERR, STORAGE } from "./constants";
import { getApiKey, getMaxTokens, getBaseUrl, getItem, setItem, removeItem } from "./storage";
import { LRUCache } from "./lru";
import { hashKey } from "./hash";
import { diagInc, diagSet, diagError, diagSuccess, diagTrackRequest, getSharedState } from "./diagnostics";

function stableStringify(value) {
  const seen = new WeakSet();
  const sorter = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

  const stringify = (v) => {
    if (v === null || v === undefined) return null;
    if (typeof v !== "object") return v;
    if (seen.has(v)) return "[Circular]";
    seen.add(v);
    if (Array.isArray(v)) return v.map(stringify);
    const out = {};
    for (const k of Object.keys(v).sort(sorter)) {
      const sv = stringify(v[k]);
      if (sv !== undefined) out[k] = sv;
    }
    return out;
  };

  return JSON.stringify(stringify(value));
}

class Semaphore {
  constructor(max) { this.max = max; this.current = 0; this.queue = []; }
  acquire() {
    if (this.current < this.max) { this.current++; return Promise.resolve(() => this._release()); }
    return new Promise((resolve) => { this.queue.push(resolve); })
      .then(() => { this.current++; return () => this._release(); });
  }
  _release() {
    this.current = Math.max(0, this.current - 1);
    const next = this.queue.shift();
    if (next) next();
  }
}

function getGlobalState() {
  const st = getSharedState() || {};
  st.memCache = st.memCache || new LRUCache(LIMITS.MEM_CACHE_ENTRIES, LIMITS.MEM_CACHE_TTL_MS);
  st.inflight = st.inflight || new Map();
  st.semaphore = st.semaphore || new Semaphore(LIMITS.MAX_CONCURRENT_REQUESTS);
  st.persistIndexLoaded = st.persistIndexLoaded || false;
  st.persistIndex = st.persistIndex || [];
  return st;
}
const ST = getGlobalState();

const PERSIST_MAX_ENTRIES = 50;

async function sleep(ms) { return new Promise((res) => setTimeout(res, ms)); }

function sanitizeCacheMode(mode) {
  return (mode === "none" || mode === "memory" || mode === "persistent") ? mode : DEFAULTS.cache;
}

function isRetriableHttpStatus(status) {
  return status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
}

function classifyHttpError(status, message) {
  if (status === 401 || status === 403) return ERR.AUTH;
  if (status === 408) return ERR.TIMEOUT;
  if (status === 429) return ERR.RATE_LIMIT;
  if (status === 413) return ERR.TOO_LARGE;
  if (status === 400 && (message || "").toLowerCase().includes("max_tokens")) return ERR.TOO_LARGE;
  return ERR.API_ERROR;
}

async function fetchWithTimeout(url, fetchOptions, timeoutMs) {
  const t = Number.isFinite(timeoutMs) ? timeoutMs : DEFAULTS.timeoutMs;

  if (typeof AbortController !== "undefined") {
    const controller = new AbortController();
    const id = setTimeout(() => controller.abort(), t);
    try {
      return await fetch(url, { ...fetchOptions, signal: controller.signal });
    } finally {
      clearTimeout(id);
    }
  }

  return await Promise.race([
    fetch(url, fetchOptions),
    new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), t))
  ]);
}

function extractChoiceText(json) {
  const first = json?.choices?.[0];
  if (!first) return { text: "", finishReason: undefined, choicesCount: 0 };

  const content = first.message?.content;
  let rendered = "";

  if (typeof content === "string") {
    rendered = content;
  } else if (Array.isArray(content)) {
    rendered = content
      .map((part) => {
        if (typeof part === "string") return part;
        if (typeof part?.text === "string") return part.text;
        if (part?.type === "text" && typeof part?.text === "string") return part.text;
        return "";
      })
      .filter(Boolean)
      .join("\n");
  }

  return { text: rendered, finishReason: first.finish_reason, choicesCount: Array.isArray(json?.choices) ? json.choices.length : 0 };
}

function isBlockedResponse(json) {
  const finish = json?.choices?.[0]?.finish_reason;
  if (!finish) return false;
  return String(finish).toLowerCase().includes("content_filter");
}

function buildDiagnostics({ json, status = 0, latencyMs = 0, cacheKey }) {
  const usage = json?.usage || {};
  const diag = {
    httpStatus: status,
    choices: Array.isArray(json?.choices) ? json.choices.length : 0,
    finishReason: json?.choices?.[0]?.finish_reason,
    cacheKey,
    latencyMs,
    usage: {
      promptTokenCount: usage.prompt_tokens,
      candidatesTokenCount: usage.completion_tokens
    },
    model: json?.model
  };
  return diag;
}

// --- Persistence Helpers ---
async function loadPersistIndex() {
  if (ST.persistIndexLoaded) return;
  try {
    const raw = await getItem(STORAGE.PERSIST_CACHE_INDEX);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) ST.persistIndex = parsed;
    }
  } catch { /* ignore */ }
  ST.persistIndexLoaded = true;
}

async function persistGet(cacheKey, ttlMs) {
  await loadPersistIndex();
  const idx = ST.persistIndex.find((e) => e && e.k === cacheKey);
  if (!idx) return null;

  try {
    const raw = await getItem("AI_PERSIST_" + cacheKey);
    if (!raw) return null;
    const obj = JSON.parse(raw);
    const t = obj?.t ? Number(obj.t) : 0;
    if (ttlMs > 0 && t && Date.now() - t > ttlMs) {
      await persistDelete(cacheKey);
      return null;
    }
    return typeof obj?.v === "string" ? obj.v : null;
  } catch {
    return null;
  }
}

async function persistSet(cacheKey, value) {
  await loadPersistIndex();
  try {
    await setItem("AI_PERSIST_" + cacheKey, JSON.stringify({ t: Date.now(), v: value }));
    ST.persistIndex = ST.persistIndex.filter((e) => e && e.k !== cacheKey);
    ST.persistIndex.push({ k: cacheKey, t: Date.now() });

    while (ST.persistIndex.length > PERSIST_MAX_ENTRIES) {
      const evict = ST.persistIndex.shift();
      if (evict?.k) await removeItem("AI_PERSIST_" + evict.k);
    }
    await setItem(STORAGE.PERSIST_CACHE_INDEX, JSON.stringify(ST.persistIndex));
  } catch { /* ignore */ }
}

async function persistDelete(cacheKey) {
  await loadPersistIndex();
  try {
    ST.persistIndex = ST.persistIndex.filter((e) => e && e.k !== cacheKey);
    await removeItem("AI_PERSIST_" + cacheKey);
    await setItem(STORAGE.PERSIST_CACHE_INDEX, JSON.stringify(ST.persistIndex));
  } catch { /* ignore */ }
}

// --- MAIN GENERATE ---
export async function openaiGenerate(req) {
  const started = Date.now();
  diagInc("requests", 1);
  diagSet("lastRequestAt", new Date().toISOString());

  const apiKey = await getApiKey();
  if (!apiKey) {
    diagError(ERR.KEY_MISSING, "API key missing");
    diagTrackRequest({ success: false, code: ERR.KEY_MISSING, message: "API key missing", latencyMs: 0 });
    return { ok: false, code: ERR.KEY_MISSING, message: "OpenAI API key missing" };
  }

  const modelRaw = (req.model || OPENAI.DEFAULT_MODEL).trim();
  const model = modelRaw.startsWith("gpt-") ? modelRaw : modelRaw;

  const cacheMode = sanitizeCacheMode(req.cache);
  const ttlMs = Math.max(0, Number(req.cacheTtlSec || DEFAULTS.cacheTtlSec)) * 1000;
  const cacheOnly = Boolean(req.cacheOnly);

  const generationConfig = req.generationConfig || {};

  let maxTokensRaw = DEFAULTS.maxTokens;
  if (typeof generationConfig.maxOutputTokens === "number") {
    maxTokensRaw = generationConfig.maxOutputTokens;
  } else {
    const stored = await getMaxTokens();
    if (stored) maxTokensRaw = stored;
  }

  let maxTokens = Math.floor(Number(maxTokensRaw));
  if (!Number.isFinite(maxTokens)) maxTokens = DEFAULTS.maxTokens;
  maxTokens = Math.min(TOKEN_LIMITS.MAX, Math.max(TOKEN_LIMITS.MIN, maxTokens));

  const messages = [];
  messages.push({ role: "system", content: String(req.system || "") });
  messages.push({ role: "user", content: String(req.user || "") });

  const responseFormat = (() => {
    if (req.responseJsonSchema) {
      const name = req.responseJsonSchema?.name || "response_schema";
      const schema = req.responseJsonSchema?.schema || req.responseJsonSchema;
      return { type: "json_schema", json_schema: { name, schema, strict: true } };
    }
    if (req.responseMimeType === "application/json") return { type: "json_object" };
    return undefined;
  })();

  const body = {
    model,
    messages,
    temperature: typeof generationConfig.temperature === "number" ? generationConfig.temperature : DEFAULTS.temperature,
    max_tokens: maxTokens
  };

  for (const k of ["topP", "topK", "stopSequences", "candidateCount"]) {
    if (generationConfig[k] !== undefined) {
      if (k === "topP") body.top_p = generationConfig[k];
    }
  }

  const filteredTools = Array.isArray(req?.tools)
    ? req.tools.filter((t) => t && typeof t === "object" && t.type === "function")
    : [];
  if (filteredTools.length > 0) body.tools = filteredTools;
  if (responseFormat) body.response_format = responseFormat;

  const rawKey = stableStringify({
    model,
    system: req.system || "",
    user: req.user || "",
    generationConfig: { ...generationConfig, maxOutputTokens: maxTokens },
    responseMimeType: req.responseMimeType || "",
    responseJsonSchema: req.responseJsonSchema || null,
    tools: body.tools || []
  });

  const cacheKey = await hashKey(rawKey);

  // CHECK CACHE
  if (cacheMode !== "none") {
    const cached = ST.memCache.get(cacheKey);
    if (typeof cached === "string") {
      diagInc("cacheHits", 1);
      const lat = Date.now() - started;
      diagSuccess({ model, latencyMs: lat, cacheKey, cached: true });
      diagTrackRequest({ success: true, code: "CACHE_MEM", model, latencyMs: lat, cached: true, functionName: req.functionName });
      return {
        ok: true,
        text: cached,
        cached: true,
        cacheKey,
        latencyMs: lat,
        diagnostics: { cacheKey, cached: true, cacheSource: "memory" }
      };
    }
    diagInc("cacheMisses", 1);
  }

  if (cacheMode === "persistent") {
    const pv = await persistGet(cacheKey, ttlMs);
    if (typeof pv === "string") {
      diagInc("cacheHits", 1);
      ST.memCache.set(cacheKey, pv);
      const lat = Date.now() - started;
      diagSuccess({ model, latencyMs: lat, cacheKey, cached: true });
      diagTrackRequest({ success: true, code: "CACHE_PERSIST", model, latencyMs: lat, cached: true, functionName: req.functionName });
      return {
        ok: true,
        text: pv,
        cached: true,
        cacheKey,
        latencyMs: lat,
        diagnostics: { cacheKey, cached: true, cacheSource: "persistent" }
      };
    }
  }

  if (cacheOnly) {
    const lat = Date.now() - started;
    const diagnostics = { cacheKey, cached: false, cacheOnly: true, cacheMode };
    diagTrackRequest({ success: false, code: ERR.CACHE_MISS, message: "Cache only mode", latencyMs: lat, model, cached: false, functionName: req.functionName });
    return { ok: false, code: ERR.CACHE_MISS, errorCode: ERR.CACHE_MISS, cacheKey, latencyMs: lat, diagnostics };
  }

  if (ST.inflight.has(cacheKey)) {
    diagInc("dedupHits", 1);
    return await ST.inflight.get(cacheKey);
  }

  // EXECUTE REQUEST
  const p = (async () => {
    const release = await ST.semaphore.acquire();
    try {
      const baseUrl = await getBaseUrl();
      const url = `${baseUrl}/chat/completions`;
      const timeoutMs = Number.isFinite(req.timeoutMs) ? req.timeoutMs : DEFAULTS.timeoutMs;
      const retries = Number.isFinite(req.retry) ? Math.max(0, Math.min(3, Math.floor(req.retry))) : DEFAULTS.retry;

      for (let attempt = 0; attempt <= retries; attempt++) {
        const attemptStart = Date.now();
        try {
          if (attempt > 0) diagInc("retries", 1);

          const resp = await fetchWithTimeout(
            url,
            {
              method: "POST",
              headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
              },
              body: JSON.stringify(body)
            },
            timeoutMs
          );

          if (!resp || !resp.ok) {
            const status = resp?.status || 0;
            let msg = `HTTP ${status}`;
            let errJson = null;
            try {
              errJson = await resp.json();
              msg = errJson?.error?.message || msg;
            } catch { /* ignore */ }

            const code = classifyHttpError(status, msg);
            const lat = Date.now() - attemptStart;

            diagTrackRequest({ success: false, code, message: msg, httpStatus: status, latencyMs: lat, model, functionName: req.functionName });

            if (attempt < retries && isRetriableHttpStatus(status)) {
              await sleep(400 * Math.pow(2, attempt));
              continue;
            }

            const diagnostics = buildDiagnostics({ json: errJson || {}, status, cacheKey, latencyMs: lat });
            diagError(code, msg, status);
            return { ok: false, code, errorCode: code, message: msg, httpStatus: status, diagnostics };
          }

          const json = await resp.json();
          const lat = Date.now() - attemptStart;
          const diagnostics = buildDiagnostics({ json, status: resp.status, cacheKey, latencyMs: lat });

          if (isBlockedResponse(json)) {
            const msg = diagnostics.finishReason ? `Blocked: ${diagnostics.finishReason}` : "Blocked by safety settings";
            diagError(ERR.BLOCKED, msg, resp.status);
            diagTrackRequest({ success: false, code: ERR.BLOCKED, message: msg, httpStatus: resp.status, latencyMs: lat, model, functionName: req.functionName });
            return { ok: false, code: ERR.BLOCKED, errorCode: ERR.BLOCKED, message: msg, httpStatus: resp.status, diagnostics };
          }

          const { text, finishReason, choicesCount } = extractChoiceText(json);
          const normalizedText = typeof text === "string" ? text : "";

          if (!normalizedText.trim()) {
            const msg = choicesCount === 0 ? "Empty response" : `Empty response (finish: ${finishReason})`;
            diagError(ERR.EMPTY_RESPONSE, msg, resp.status);
            diagTrackRequest({ success: false, code: ERR.EMPTY_RESPONSE, message: msg, httpStatus: resp.status, latencyMs: lat, model, functionName: req.functionName });
            return { ok: false, code: ERR.EMPTY_RESPONSE, errorCode: ERR.EMPTY_RESPONSE, message: msg, httpStatus: resp.status, diagnostics };
          }

          let cleaned = normalizedText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();

          if (cacheMode !== "none") ST.memCache.set(cacheKey, cleaned);
          if (cacheMode === "persistent") await persistSet(cacheKey, cleaned);

          diagSuccess({ model, latencyMs: lat, cacheKey, cached: false });
          diagTrackRequest({
            success: true,
            code: "OK",
            message: req.user || "",
            usage: diagnostics.usage,
            latencyMs: lat,
            model,
            cached: false,
            functionName: req.functionName
          });

          return { ok: true, text: cleaned, cached: false, cacheKey, latencyMs: lat, diagnostics };
        } catch (e) {
          const msg = (e?.name === "AbortError" || e?.message === "timeout") ? "Timeout" : (e?.message || "Network error");
          const code = msg === "Timeout" ? ERR.TIMEOUT : ERR.API_ERROR;
          const lat = Date.now() - attemptStart;

          if (attempt < retries) {
            await sleep(300 * Math.pow(2, attempt));
            continue;
          }

          diagError(code, msg, 0);
          diagTrackRequest({ success: false, code, message: msg, latencyMs: lat, model, functionName: req.functionName });

          const diagnostics = buildDiagnostics({ json: {}, status: 0, cacheKey, latencyMs: lat });
          return { ok: false, code, errorCode: code, message: msg, httpStatus: 0, diagnostics };
        }
      }

      diagError(ERR.API_ERROR, "Unknown error", 0);
      diagTrackRequest({ success: false, code: ERR.API_ERROR, message: "Unknown", latencyMs: Date.now() - started, model, functionName: req.functionName });
      return { ok: false, code: ERR.API_ERROR, errorCode: ERR.API_ERROR, message: "Unknown error", httpStatus: 0, diagnostics: {} };
    } finally {
      release();
    }
  })();

  ST.inflight.set(cacheKey, p);
  try { return await p; }
  finally { ST.inflight.delete(cacheKey); }
}

export async function openaiMinimalTest(options = {}) {
  const res = await openaiGenerate({
    model: options.model,
    system: "You are a connectivity test. Reply with exactly: OK",
    user: "Return OK.",
    generationConfig: { temperature: 0.0, maxOutputTokens: 256 },
    responseMimeType: "text/plain",
    cache: "none",
    timeoutMs: options.timeoutMs,
    retry: 0
  });

  if (!res.ok) return res;
  return { ok: true, text: "OK", diagnostics: res.diagnostics };
}
