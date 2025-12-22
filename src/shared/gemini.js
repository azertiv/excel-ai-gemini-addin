import { GEMINI, DEFAULTS, LIMITS, ERR, STORAGE } from "./constants";
import { getApiKey, getItem, setItem, removeItem } from "./storage";
import { LRUCache } from "./lru";
import { hashKey } from "./hash";
import { diagInc, diagSet, diagError, diagSuccess, getSharedState } from "./diagnostics";

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

function classifyHttpError(status) {
  if (status === 401 || status === 403) return ERR.AUTH;
  if (status === 408) return ERR.TIMEOUT;
  if (status === 429) return ERR.RATE_LIMIT;
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

function extractCandidateText(json) {
  const candidates = json?.candidates;
  if (!Array.isArray(candidates) || candidates.length === 0) return "";
  const parts = candidates[0]?.content?.parts;
  if (!Array.isArray(parts) || parts.length === 0) return "";

  const rendered = parts
    .map((p) => {
      if (typeof p?.text === "string") return p.text;
      if (p?.functionCall) {
        const name = p.functionCall.name || "functionCall";
        const args = p.functionCall.args ? `(${JSON.stringify(p.functionCall.args)})` : "";
        return `${name}${args}`;
      }
      if (p?.inlineData?.data) return p.inlineData.data;
      if (p?.executableCode?.code) return p.executableCode.code;
      return "";
    })
    .filter(Boolean)
    .join("\n");

  return rendered;
}

function isBlockedResponse(json) {
  if (json?.promptFeedback?.blockReason) return true;
  const finish = json?.candidates?.[0]?.finishReason;
  if (finish && String(finish).toUpperCase().includes("SAFETY")) return true;
  return false;
}

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

export async function geminiGenerate(req) {
  const started = Date.now();
  diagInc("requests", 1);
  diagSet("lastRequestAt", new Date().toISOString());

  const apiKey = await getApiKey();
  if (!apiKey) {
    diagError(ERR.KEY_MISSING, "API key missing");
    return { ok: false, code: ERR.KEY_MISSING, message: "Gemini API key missing" };
  }

  const modelRaw = (req.model || GEMINI.DEFAULT_MODEL).trim();
  const model = modelRaw.startsWith("models/") ? modelRaw.slice("models/".length) : modelRaw;

  const cacheMode = sanitizeCacheMode(req.cache);
  const ttlMs = Math.max(0, Number(req.cacheTtlSec || DEFAULTS.cacheTtlSec)) * 1000;

  const generationConfig = req.generationConfig || {};
  const body = {
    systemInstruction: { role: "system", parts: [{ text: String(req.system || "") }] },
    contents: [{ role: "user", parts: [{ text: String(req.user || "") }] }],
    generationConfig: {
      temperature: typeof generationConfig.temperature === "number" ? generationConfig.temperature : DEFAULTS.temperature,
      maxOutputTokens: typeof generationConfig.maxOutputTokens === "number" ? generationConfig.maxOutputTokens : DEFAULTS.maxTokens
    }
  };

  for (const k of ["topP", "topK", "candidateCount", "stopSequences"]) {
    if (generationConfig[k] !== undefined) body.generationConfig[k] = generationConfig[k];
  }

  if (req.responseMimeType) body.generationConfig.responseMimeType = req.responseMimeType;
  if (req.responseJsonSchema) body.generationConfig.responseJsonSchema = req.responseJsonSchema;

  const rawKey = stableStringify({
    model,
    system: req.system || "",
    user: req.user || "",
    generationConfig: body.generationConfig,
    responseMimeType: req.responseMimeType || "",
    responseJsonSchema: req.responseJsonSchema || null
  });

  const cacheKey = await hashKey(rawKey);

  if (cacheMode !== "none") {
    const cached = ST.memCache.get(cacheKey);
    if (typeof cached === "string") {
      diagInc("cacheHits", 1);
      diagSuccess({ model, latencyMs: Date.now() - started, cacheKey, cached: true });
      return { ok: true, text: cached, cached: true, cacheKey, latencyMs: Date.now() - started };
    }
    diagInc("cacheMisses", 1);
  }

  if (cacheMode === "persistent") {
    const pv = await persistGet(cacheKey, ttlMs);
    if (typeof pv === "string") {
      diagInc("cacheHits", 1);
      ST.memCache.set(cacheKey, pv);
      diagSuccess({ model, latencyMs: Date.now() - started, cacheKey, cached: true });
      return { ok: true, text: pv, cached: true, cacheKey, latencyMs: Date.now() - started };
    }
  }

  if (ST.inflight.has(cacheKey)) {
    diagInc("dedupHits", 1);
    return await ST.inflight.get(cacheKey);
  }

  const p = (async () => {
    const release = await ST.semaphore.acquire();
    try {
      const url = `${GEMINI.BASE_URL}/models/${encodeURIComponent(model)}:generateContent`;
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
                "x-goog-api-key": apiKey
              },
              body: JSON.stringify(body)
            },
            timeoutMs
          );

          if (!resp || !resp.ok) {
            const status = resp?.status || 0;
            let msg = `HTTP ${status}`;
            try {
              const errJson = await resp.json();
              msg = errJson?.error?.message || msg;
            } catch { /* ignore */ }

            const code = classifyHttpError(status);

            if (attempt < retries && isRetriableHttpStatus(status)) {
              await sleep(400 * Math.pow(2, attempt));
              continue;
            }

            diagError(code, msg, status);
            return { ok: false, code, message: msg, httpStatus: status };
          }

          const json = await resp.json();

          if (isBlockedResponse(json)) {
            diagError(ERR.BLOCKED, "Blocked by safety settings", resp.status);
            return { ok: false, code: ERR.BLOCKED, message: "Blocked by safety settings", httpStatus: resp.status };
          }

          let text = extractCandidateText(json);
          text = typeof text === "string" ? text : "";

          if (!text.trim()) {
            diagError(ERR.EMPTY_RESPONSE, "Empty response", resp.status);
            return { ok: false, code: ERR.EMPTY_RESPONSE, message: "Empty response", httpStatus: resp.status };
          }

          text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
          if (text.length > LIMITS.MAX_CELL_CHARS) text = text.slice(0, LIMITS.MAX_CELL_CHARS) + "\n…(truncated)";

          if (cacheMode !== "none") ST.memCache.set(cacheKey, text);
          if (cacheMode === "persistent") await persistSet(cacheKey, text);

          diagSuccess({ model, latencyMs: Date.now() - attemptStart, cacheKey, cached: false });
          return { ok: true, text, cached: false, cacheKey, latencyMs: Date.now() - attemptStart };
        } catch (e) {
          const msg = (e?.name === "AbortError" || e?.message === "timeout") ? "Timeout" : (e?.message || "Network error");
          const code = msg === "Timeout" ? ERR.TIMEOUT : ERR.API_ERROR;

          if (attempt < retries) {
            await sleep(300 * Math.pow(2, attempt));
            continue;
          }

          diagError(code, msg, 0);
          return { ok: false, code, message: msg, httpStatus: 0 };
        }
      }

      diagError(ERR.API_ERROR, "Unknown error", 0);
      return { ok: false, code: ERR.API_ERROR, message: "Unknown error", httpStatus: 0 };
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
    model: options.model,
    system: "You are a connectivity test. Reply with exactly: OK",
    user: "Return OK.",
    generationConfig: { temperature: 0.0, maxOutputTokens: 8 },
    responseMimeType: "text/plain",
    cache: "none",
    timeoutMs: options.timeoutMs,
    retry: 0
  });

  if (!res.ok) return res;
  return { ok: true, text: "OK" };
}
