function getGlobalState() {
  if (typeof window === "undefined") return {};
  window.__AI_ADDIN_STATE__ = window.__AI_ADDIN_STATE__ || {};
  return window.__AI_ADDIN_STATE__;
}

const state = getGlobalState();

function initDiagnostics() {
  if (state.diagnostics) return;
  state.diagnostics = {
    startedAt: new Date().toISOString(),
    backend: "",
    requests: 0,
    success: 0,
    failures: 0,
    retries: 0,
    cacheHits: 0,
    cacheMisses: 0,
    dedupHits: 0,
    lastRequestAt: "",
    lastSuccessAt: "",
    lastErrorAt: "",
    lastErrorCode: "",
    lastErrorMessage: "",
    lastHttpStatus: 0,
    lastModel: "",
    lastLatencyMs: 0,
    lastCacheKey: "",
    events: []
  };
}
initDiagnostics();

export function diagSet(k, v) { state.diagnostics[k] = v; }
export function diagInc(k, n = 1) { state.diagnostics[k] = (state.diagnostics[k] || 0) + n; }

export function diagError(code, message, httpStatus = 0) {
  state.diagnostics.lastErrorAt = new Date().toISOString();
  state.diagnostics.lastErrorCode = code || "";
  state.diagnostics.lastErrorMessage = message || "";
  state.diagnostics.lastHttpStatus = httpStatus || 0;
  diagInc("failures", 1);
  pushEvent("error", { code, message, httpStatus });
}

export function diagSuccess({ model, latencyMs, cacheKey, cached } = {}) {
  state.diagnostics.lastSuccessAt = new Date().toISOString();
  state.diagnostics.lastModel = model || state.diagnostics.lastModel;
  state.diagnostics.lastLatencyMs = Number.isFinite(latencyMs) ? latencyMs : state.diagnostics.lastLatencyMs;
  state.diagnostics.lastCacheKey = cacheKey || state.diagnostics.lastCacheKey;
  diagInc("success", 1);
  pushEvent("success", { model, latencyMs, cached: !!cached });
}

export function pushEvent(kind, payload) {
  const ev = { at: new Date().toISOString(), kind, payload: payload || {} };
  state.diagnostics.events.push(ev);
  if (state.diagnostics.events.length > 20) {
    state.diagnostics.events.splice(0, state.diagnostics.events.length - 20);
  }
}

export function getDiagnosticsSnapshot() {
  return JSON.parse(JSON.stringify(state.diagnostics || {}));
}

export function formatDiagnosticsForUi(snapshot) {
  if (!snapshot) return "";
  const lines = [];
  lines.push(`Started: ${snapshot.startedAt || ""}`);
  if (snapshot.backend) lines.push(`Storage backend: ${snapshot.backend}`);
  lines.push(`Requests: ${snapshot.requests || 0} (ok: ${snapshot.success || 0}, fail: ${snapshot.failures || 0})`);
  lines.push(`Cache: hits ${snapshot.cacheHits || 0}, misses ${snapshot.cacheMisses || 0}, dedup ${snapshot.dedupHits || 0}`);
  lines.push(`Retries: ${snapshot.retries || 0}`);
  if (snapshot.lastRequestAt) lines.push(`Last request: ${snapshot.lastRequestAt} (model: ${snapshot.lastModel || "?"}, ${snapshot.lastLatencyMs || 0} ms)`);
  if (snapshot.lastErrorAt) {
    lines.push(`Last error: ${snapshot.lastErrorAt}`);
    lines.push(`  ${snapshot.lastErrorCode || ""}${snapshot.lastHttpStatus ? " (HTTP " + snapshot.lastHttpStatus + ")" : ""}`);
    if (snapshot.lastErrorMessage) lines.push(`  ${snapshot.lastErrorMessage}`);
  }
  if (Array.isArray(snapshot.events) && snapshot.events.length) {
    lines.push(`Events (last ${Math.min(20, snapshot.events.length)}):`);
    for (const e of snapshot.events.slice(-10)) {
      lines.push(`  - ${e.at} ${e.kind}${e.payload?.cached ? " [cached]" : ""}`);
    }
  }
  return lines.join("\n");
}

export function getSharedState() {
  return state;
}
