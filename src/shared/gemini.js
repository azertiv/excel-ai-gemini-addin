// src/shared/gemini.js

import { GEMINI, OPENAI, PROVIDERS, DEFAULTS, LIMITS, TOKEN_LIMITS, ERR, STORAGE } from "./constants";
import { getApiKey, getMaxTokens, getItem, setItem, removeItem, getProvider, getModel } from "./storage";
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

function normalizeProvider(p) {
  return p === PROVIDERS.OPENAI ? PROVIDERS.OPENAI : PROVIDERS.GEMINI;
}

function stripModelPrefix(model) {
  if (!model) return "";
  const m = String(model).trim();
  return m.startsWith("models/") ? m.slice("models/".length) : m;
}

async function resolveModel(provider, requestedModel) {
  const raw = stripModelPrefix(requestedModel);
  if (raw) return raw;

  const stored = await getModel(provider);
  if (stored) return stripModelPrefix(stored);

  return provider === PROVIDERS.OPENAI ? OPENAI.DEFAULT_MODEL : stripModelPrefix(GEMINI.DEFAULT_MODEL);
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

function looksTooLarge(message) {
  const m = String(message || "").toLowerCase();
  // Common patterns returned by Google APIs when a request exceeds payload/context limits.
  return (
    m.includes("too large") ||
    m.includes("exceeds") ||
    m.includes("exceed") ||
    m.includes("maximum") && m.includes("token") ||
    m.includes("context") && m.includes("limit") ||
    m.includes("request payload") ||
    m.includes("payload size")
  );
}

function classifyHttpError(status, message) {
  if (status === 401 || status === 403) return ERR.AUTH;
  if (status === 408) return ERR.TIMEOUT;
  if (status === 429) return ERR.RATE_LIMIT;
  if (status === 413) return ERR.TOO_LARGE;
  if (status === 400 && looksTooLarge(message)) return ERR.TOO_LARGE;
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

function normalizeUsage(provider, json) {
  if (provider === PROVIDERS.OPENAI) {
    const u = json?.usage;
    if (!u) return undefined;
    const prompt = Number(u.prompt_tokens) || 0;
    const completion = Number(u.completion_tokens) || 0;
    const total = Number(u.total_tokens) || (prompt + completion);
    return { promptTokenCount: prompt, candidatesTokenCount: completion, totalTokenCount: total };
  }

  const u = json?.usageMetadata;
  if (!u) return undefined;
  const prompt = Number(u.promptTokenCount) || 0;
  const completion = Number(u.candidatesTokenCount) || 0;
  const total = Number(u.totalTokenCount) || (prompt + completion);
  return { promptTokenCount: prompt, candidatesTokenCount: completion, totalTokenCount: total };
}

function extractGeminiText(json) {
  const candidates = json?.candidates;
  if (!Array.isArray(candidates) || candidates.length === 0) return { text: "", candidatesCount: 0 };

  const first = candidates[0];
  const content = first?.content;
  const parts = Array.isArray(content)
    ? content.flatMap((c) => (Array.isArray(c?.parts) ? c.parts : [])).filter(Boolean)
    : (Array.isArray(content?.parts) ? content.parts : []);

  if (!Array.isArray(parts) || parts.length === 0) {
    return { text: "", candidatesCount: candidates.length, finishReason: first?.finishReason };
  }

  const rendered = parts
    .map((p) => {
      if (typeof p?.text === "string") return p.text;
      if (p?.functionCall) {
        const name = p.functionCall.name || "functionCall";
        const args = p.functionCall.args ? `(${JSON.stringify(p.functionCall.args)})` : "";
        return `${name}${args}`;
      }
      return "";
    })
    .filter(Boolean)
    .join("\n");

  return { text: rendered, candidatesCount: candidates.length, finishReason: first?.finishReason };
}

function extractOpenAIText(json) {
  const choices = json?.choices;
  if (!Array.isArray(choices) || choices.length === 0) return { text: "", candidatesCount: 0 };

  const first = choices[0];
  const finishReason = first?.finish_reason || first?.finishReason;
  const message = first?.message || {};
  const parts = [];

  const content = message.content;
  if (typeof content === "string") {
    parts.push(content);
  } else if (Array.isArray(content)) {
    for (const c of content) {
      if (typeof c?.text === "string") parts.push(c.text);
      else if (typeof c?.content === "string") parts.push(c.content);
      else if (typeof c?.value === "string") parts.push(c.value);
    }
  }

  if (Array.isArray(message.tool_calls)) {
    for (const tc of message.tool_calls) {
      const name = tc?.function?.name || tc?.type || "tool";
      const args = tc?.function?.arguments || "";
      parts.push(`${name}(${args})`);
    }
  }

  const rendered = parts.filter(Boolean).join("\n");
  return { text: rendered, candidatesCount: choices.length, finishReason };
}

function isBlockedResponse(provider, json) {
  if (provider === PROVIDERS.OPENAI) {
    const finish = json?.choices?.[0]?.finish_reason || json?.choices?.[0]?.finishReason;
    if (!finish) return false;
    return String(finish).toLowerCase().includes("content_filter");
  }

  if (json?.promptFeedback?.blockReason) return true;
  const finish = json?.candidates?.[0]?.finishReason;
  if (finish && String(finish).toUpperCase().includes("SAFETY")) return true;
  return false;
}

function normalizeToolsForOpenAI(tools) {
  if (!Array.isArray(tools)) return [];
  const out = [];
  for (const t of tools) {
    if (!t) continue;
    if (t.type === "function" && t.function && t.function.name) {
      out.push({ type: "function", function: t.function });
      continue;
    }
    if (Array.isArray(t.functionDeclarations)) {
      for (const fn of t.functionDeclarations) {
        if (fn && fn.name) out.push({ type: "function", function: { name: fn.name, description: fn.description, parameters: fn.parameters } });
      }
    }
  }
  return out;
}

function buildDiagnostics({ provider, json, status = 0, latencyMs = 0, cacheKey }) {
  const candidates = provider === PROVIDERS.OPENAI
    ? (Array.isArray(json?.choices) ? json.choices.length : 0)
    : (Array.isArray(json?.candidates) ? json.candidates.length : 0);

  const finishReason = provider === PROVIDERS.OPENAI
    ? (json?.choices?.[0]?.finish_reason || json?.choices?.[0]?.finishReason)
    : json?.candidates?.[0]?.finishReason;

  const blockReason = provider === PROVIDERS.OPENAI
    ? (finishReason === "content_filter" ? "content_filter" : undefined)
    : json?.promptFeedback?.blockReason;

  const diag = {
    provider,
    httpStatus: status,
    candidates,
    finishReason,
    blockReason,
    safety: provider === PROVIDERS.OPENAI ? undefined : json?.candidates?.[0]?.safetyRatings,
    modelVersion: json?.modelVersion || json?.model,
    usage: normalizeUsage(provider, json),
    cacheKey,
    latencyMs
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
export async function geminiGenerate(req) {
  const started = Date.now();
  diagInc("requests", 1);
  diagSet("lastRequestAt", new Date().toISOString());

  const provider = normalizeProvider(req?.provider || (await getProvider()));
  const model = await resolveModel(provider, req.model);

  const apiKey = await getApiKey(provider);
  if (!apiKey) {
    const msg = `${provider === PROVIDERS.OPENAI ? "OpenAI" : "Gemini"} API key missing`;
    diagError(ERR.KEY_MISSING, msg, 0, provider);
    diagTrackRequest({ success: false, code: ERR.KEY_MISSING, message: msg, latencyMs: 0, provider, model, functionName: req.functionName });
    return { ok: false, code: ERR.KEY_MISSING, message: msg, provider, model };
  }

  const cacheMode = sanitizeCacheMode(req.cache);
  const ttlMs = Math.max(0, Number(req.cacheTtlSec || DEFAULTS.cacheTtlSec)) * 1000;
  const cacheOnly = Boolean(req.cacheOnly);

  const generationConfig = req.generationConfig || {};

  // Logic priority: req.generationConfig.maxOutputTokens > storage setting > DEFAULTS.maxTokens
  // Clamp to the UI bounds (32..128000) to avoid silent provider errors.
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

  const systemText = String(req.system || "");
  const userText = String(req.user || "");
  const temp = typeof generationConfig.temperature === "number" ? generationConfig.temperature : DEFAULTS.temperature;

  const geminiBody = {
    systemInstruction: { role: "system", parts: [{ text: systemText }] },
    contents: [{ role: "user", parts: [{ text: userText }] }],
    generationConfig: {
      temperature: temp,
      maxOutputTokens: maxTokens
    }
  };

  const openaiBody = {
    model,
    messages: [],
    temperature: temp
  };
  if (systemText) openaiBody.messages.push({ role: "system", content: systemText });
  openaiBody.messages.push({ role: "user", content: userText });
  openaiBody.max_tokens = maxTokens;

  if (Array.isArray(req?.tools) && req.tools.length > 0) {
    if (provider === PROVIDERS.OPENAI) {
      const converted = normalizeToolsForOpenAI(req.tools);
      if (converted.length > 0) openaiBody.tools = converted;
    } else {
      geminiBody.tools = req.tools;
    }
  }

  for (const k of ["topP", "topK", "candidateCount", "stopSequences"]) {
    if (generationConfig[k] !== undefined) geminiBody.generationConfig[k] = generationConfig[k];
  }
  if (generationConfig.topP !== undefined) openaiBody.top_p = generationConfig.topP;
  if (generationConfig.stopSequences !== undefined) openaiBody.stop = generationConfig.stopSequences;

  if (req.responseMimeType) geminiBody.generationConfig.responseMimeType = req.responseMimeType;
  if (req.responseJsonSchema) geminiBody.generationConfig.responseJsonSchema = req.responseJsonSchema;

  if (req.responseJsonSchema) {
    openaiBody.response_format = { type: "json_schema", json_schema: { name: "schema", schema: req.responseJsonSchema, strict: true } };
  } else if (req.responseMimeType === "application/json") {
    openaiBody.response_format = { type: "json_object" };
  }

  const rawKey = stableStringify({
    provider,
    model,
    system: systemText,
    user: userText,
    generationConfig: geminiBody.generationConfig,
    responseMimeType: req.responseMimeType || "",
    responseJsonSchema: req.responseJsonSchema || null,
    tools: req.tools || []
  });

  const cacheKey = await hashKey(rawKey);

  // CHECK CACHE
  if (cacheMode !== "none") {
    const cached = ST.memCache.get(cacheKey);
    if (typeof cached === "string") {
      diagInc("cacheHits", 1);
      const lat = Date.now() - started;
      diagSuccess({ model, latencyMs: lat, cacheKey, cached: true, provider });
      diagTrackRequest({ success: true, code: "CACHE_MEM", model, latencyMs: lat, cached: true, functionName: req.functionName, provider });
      return {
        ok: true,
        text: cached,
        cached: true,
        provider,
        model,
        cacheKey,
        latencyMs: lat,
        diagnostics: { cacheKey, cached: true, cacheSource: "memory", provider }
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
      diagSuccess({ model, latencyMs: lat, cacheKey, cached: true, provider });
      diagTrackRequest({ success: true, code: "CACHE_PERSIST", model, latencyMs: lat, cached: true, functionName: req.functionName, provider });
      return {
        ok: true,
        text: pv,
        cached: true,
        provider,
        model,
        cacheKey,
        latencyMs: lat,
        diagnostics: { cacheKey, cached: true, cacheSource: "persistent", provider }
      };
    }
  }

  if (cacheOnly) {
    const lat = Date.now() - started;
    const diagnostics = { cacheKey, cached: false, cacheOnly: true, cacheMode, provider };
    diagTrackRequest({ success: false, code: ERR.CACHE_MISS, message: "Cache only mode", latencyMs: lat, model, cached: false, functionName: req.functionName, provider });
    return { ok: false, code: ERR.CACHE_MISS, errorCode: ERR.CACHE_MISS, cacheKey, latencyMs: lat, provider, diagnostics };
  }

  if (ST.inflight.has(cacheKey)) {
    diagInc("dedupHits", 1);
    return await ST.inflight.get(cacheKey);
  }

  // EXECUTE REQUEST
  const p = (async () => {
    const release = await ST.semaphore.acquire();
    try {
      const timeoutMs = Number.isFinite(req.timeoutMs) ? req.timeoutMs : DEFAULTS.timeoutMs;
      const retries = Number.isFinite(req.retry) ? Math.max(0, Math.min(3, Math.floor(req.retry))) : DEFAULTS.retry;

      const url = provider === PROVIDERS.OPENAI
        ? `${OPENAI.BASE_URL}/chat/completions`
        : `${GEMINI.BASE_URL}/models/${encodeURIComponent(model)}:generateContent`;

      const fetchOptions = provider === PROVIDERS.OPENAI
        ? {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "Authorization": `Bearer ${apiKey}`
            },
            body: JSON.stringify(openaiBody)
          }
        : {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "x-goog-api-key": apiKey
            },
            body: JSON.stringify(geminiBody)
          };

      for (let attempt = 0; attempt <= retries; attempt++) {
        const attemptStart = Date.now();
        try {
          if (attempt > 0) diagInc("retries", 1);

          const resp = await fetchWithTimeout(url, fetchOptions, timeoutMs);

          if (!resp || !resp.ok) {
            const status = resp?.status || 0;
            let msg = `HTTP ${status}`;
            let errJson = null;
            try {
              errJson = await resp.json();
              msg = errJson?.error?.message || errJson?.message || msg;
            } catch { /* ignore */ }

            const code = classifyHttpError(status, msg);
            const lat = Date.now() - attemptStart;

            diagTrackRequest({ success: false, code, message: msg, httpStatus: status, latencyMs: lat, model, functionName: req.functionName, provider });

            if (attempt < retries && isRetriableHttpStatus(status)) {
              await sleep(400 * Math.pow(2, attempt));
              continue;
            }

            const diagnostics = buildDiagnostics({ provider, json: errJson || {}, status, cacheKey, latencyMs: lat });
            diagError(code, msg, status, provider);
            return { ok: false, code, errorCode: code, message: msg, httpStatus: status, diagnostics, provider, model };
          }

          // SUCCES
          const json = await resp.json();
          const lat = Date.now() - attemptStart;
          const diagnostics = buildDiagnostics({ provider, json, status: resp.status, cacheKey, latencyMs: lat });

          if (isBlockedResponse(provider, json)) {
            const msg = diagnostics.blockReason ? `Blocked: ${diagnostics.blockReason}` : "Blocked by safety settings";
            diagError(ERR.BLOCKED, msg, resp.status, provider);
            diagTrackRequest({ success: false, code: ERR.BLOCKED, message: msg, httpStatus: resp.status, latencyMs: lat, model, functionName: req.functionName, provider });
            return { ok: false, code: ERR.BLOCKED, errorCode: ERR.BLOCKED, message: msg, httpStatus: resp.status, diagnostics, provider, model };
          }

          const { text, candidatesCount, finishReason } = provider === PROVIDERS.OPENAI ? extractOpenAIText(json) : extractGeminiText(json);
          const normalizedText = typeof text === "string" ? text : "";

          if (!normalizedText.trim()) {
            const msg = candidatesCount === 0 ? "Empty response" : `Empty response (finish: ${finishReason})`;
            diagError(ERR.EMPTY_RESPONSE, msg, resp.status, provider);
            diagTrackRequest({ success: false, code: ERR.EMPTY_RESPONSE, message: msg, httpStatus: resp.status, latencyMs: lat, model, functionName: req.functionName, provider });
            return { ok: false, code: ERR.EMPTY_RESPONSE, errorCode: ERR.EMPTY_RESPONSE, message: msg, httpStatus: resp.status, diagnostics, provider, model };
          }

          // IMPORTANT: do not truncate the raw model output here.
          // Many Excel functions expect to parse JSON/TSV returned by the model; truncation would corrupt it.
          // Cell-length constraints are enforced later (when returning a single-cell string result).
          let cleaned = normalizedText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();

          if (cacheMode !== "none") ST.memCache.set(cacheKey, cleaned);
          if (cacheMode === "persistent") await persistSet(cacheKey, cleaned);

          const groundingMetadata = provider === PROVIDERS.GEMINI
            ? (json?.candidates?.[0]?.groundingMetadata || json?.groundingMetadata)
            : undefined;

          diagSuccess({ model, latencyMs: lat, cacheKey, cached: false, provider });
          diagTrackRequest({ 
            success: true, 
            code: "OK", 
            message: req.user || "", // Log prompt as message for visibility
            usage: diagnostics.usage, // Important pour le compteur de tokens
            latencyMs: lat, 
            model, 
            cached: false,
            functionName: req.functionName,
            provider
          });

          return { ok: true, text: cleaned, cached: false, provider, model, cacheKey, latencyMs: lat, diagnostics, groundingMetadata };
          
        } catch (e) {
          const msg = (e?.name === "AbortError" || e?.message === "timeout") ? "Timeout" : (e?.message || "Network error");
          const code = msg === "Timeout" ? ERR.TIMEOUT : ERR.API_ERROR;
          const lat = Date.now() - attemptStart;

          if (attempt < retries) {
            await sleep(300 * Math.pow(2, attempt));
            continue;
          }

          diagError(code, msg, 0, provider);
          diagTrackRequest({ success: false, code, message: msg, latencyMs: lat, model, functionName: req.functionName, provider });
          
          const diagnostics = buildDiagnostics({ provider, json: {}, status: 0, cacheKey, latencyMs: lat });
          return { ok: false, code, errorCode: code, message: msg, httpStatus: 0, diagnostics, provider, model };
        }
      }

      diagError(ERR.API_ERROR, "Unknown error", 0, provider);
      diagTrackRequest({ success: false, code: ERR.API_ERROR, message: "Unknown", latencyMs: Date.now() - started, model, functionName: req.functionName, provider });
      return { ok: false, code: ERR.API_ERROR, errorCode: ERR.API_ERROR, message: "Unknown error", httpStatus: 0, diagnostics: {}, provider, model };
    } finally {
      release();
    }
  })();

  ST.inflight.set(cacheKey, p);
  try { return await p; }
  finally { ST.inflight.delete(cacheKey); }
}

export async function geminiMinimalTest(options = {}) {
  const res = await geminiGenerate({
    provider: options.provider,
    model: options.model,
    system: "You are a connectivity test. Reply with exactly: OK",
    user: "Return OK.",
    generationConfig: { temperature: 0.0, maxOutputTokens: 1024 },
    responseMimeType: "text/plain",
    cache: "none",
    timeoutMs: options.timeoutMs,
    retry: 0
  });

  if (!res.ok) return res;
  return { ok: true, text: "OK", diagnostics: res.diagnostics, provider: res.provider, model: res.model };
}
