// src/shared/diagnostics.js

import { PROVIDERS } from "./constants";

// Tarifs approximatifs par million de tokens (à ajuster si besoin)
// Valeurs par défaut: Gemini 3.0 Flash (~0.10 / 0.40), OpenAI Mini (~0.15 / 0.60).
const COSTS = {
  [PROVIDERS.GEMINI]: { in: 0.10, out: 0.40 },
  [PROVIDERS.OPENAI]: { in: 0.15, out: 0.60 }
};
const DEFAULT_PROVIDER = PROVIDERS.GEMINI;

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
    // Compteurs globaux
    requests: 0,
    success: 0,
    failures: 0,
    retries: 0,
    cacheHits: 0,
    cacheMisses: 0,
    dedupHits: 0,
    
    // Stats Tokens & Coûts
    totalInputTokens: 0,
    totalOutputTokens: 0,
    estimatedCostUSD: 0,
    
    // Derniers états
    lastRequestAt: "",
    lastSuccessAt: "",
    lastErrorAt: "",
    lastErrorCode: "",
    lastErrorMessage: "",
    lastModel: "",
    lastProvider: "",
    lastLatencyMs: 0,
    
    // Logs détaillés (les 50 derniers)
    logs: [] 
  };
}
initDiagnostics();

export function diagSet(k, v) { state.diagnostics[k] = v; }
export function diagInc(k, n = 1) { state.diagnostics[k] = (state.diagnostics[k] || 0) + n; }

export function diagError(code, message, httpStatus = 0, provider = DEFAULT_PROVIDER) {
  state.diagnostics.lastErrorAt = new Date().toISOString();
  state.diagnostics.lastErrorCode = code || "";
  state.diagnostics.lastErrorMessage = message || "";
  state.diagnostics.lastProvider = provider || DEFAULT_PROVIDER;
  diagInc("failures", 1);
  // On loggue aussi l'erreur dans l'historique détaillé
  diagTrackRequest({ success: false, code, message, httpStatus, provider });
}

export function diagSuccess({ model, latencyMs, cacheKey, cached, provider } = {}) {
  state.diagnostics.lastSuccessAt = new Date().toISOString();
  state.diagnostics.lastModel = model || state.diagnostics.lastModel;
  state.diagnostics.lastProvider = provider || state.diagnostics.lastProvider;
  state.diagnostics.lastLatencyMs = Number.isFinite(latencyMs) ? latencyMs : state.diagnostics.lastLatencyMs;
  diagInc("success", 1);
}

/**
 * Enregistre une requête terminée dans l'historique et met à jour les coûts.
 */
export function diagTrackRequest({ success, code, message, usage, latencyMs, model, cached, httpStatus, functionName, provider = DEFAULT_PROVIDER }) {
  // Mise à jour des compteurs basiques si pas déjà fait par diagSuccess/diagError
  // (Note: diagSuccess/Error incrémentent déjà success/failures, ici on gère logs et coûts)

  // Calcul Tokens & Coûts (uniquement si pas en cache)
  const input = usage?.promptTokenCount || 0;
  const output = usage?.candidatesTokenCount || 0;
  
  if (!cached && success) {
    state.diagnostics.totalInputTokens = (state.diagnostics.totalInputTokens || 0) + input;
    state.diagnostics.totalOutputTokens = (state.diagnostics.totalOutputTokens || 0) + output;
    
    const costCfg = COSTS[provider] || COSTS[DEFAULT_PROVIDER];
    const cost = (input / 1_000_000 * costCfg.in) + (output / 1_000_000 * costCfg.out);
    state.diagnostics.estimatedCostUSD = (state.diagnostics.estimatedCostUSD || 0) + cost;
  }

  // Création de l'entrée de log
  const entry = {
    id: Date.now().toString(36) + Math.random().toString(36).substr(2, 5),
    at: new Date().toISOString(),
    success,
    code: code || (success ? "OK" : "ERR"),
    functionName: functionName || "?",
    message: message || "",
    model: model || "?",
    provider: provider || DEFAULT_PROVIDER,
    latencyMs: latencyMs || 0,
    inputTokens: input,
    outputTokens: output,
    cached: !!cached,
    httpStatus: httpStatus || 0
  };

  // Gestion du buffer circulaire (50 derniers logs)
  if (!state.diagnostics.logs) state.diagnostics.logs = [];
  state.diagnostics.logs.unshift(entry);
  if (state.diagnostics.logs.length > 50) state.diagnostics.logs.pop();
}

export function resetDiagnosticsLogs() {
  if (!state.diagnostics) return;

  state.diagnostics.logs = [];
  state.diagnostics.totalInputTokens = 0;
  state.diagnostics.totalOutputTokens = 0;
  state.diagnostics.estimatedCostUSD = 0;
  state.diagnostics.requests = 0;
  state.diagnostics.success = 0;
  state.diagnostics.failures = 0;
  state.diagnostics.retries = 0;
  state.diagnostics.cacheHits = 0;
  state.diagnostics.cacheMisses = 0;
  state.diagnostics.dedupHits = 0;
  state.diagnostics.startedAt = new Date().toISOString();
  state.diagnostics.lastRequestAt = "";
  state.diagnostics.lastSuccessAt = "";
  state.diagnostics.lastErrorAt = "";
  state.diagnostics.lastErrorCode = "";
  state.diagnostics.lastErrorMessage = "";
  state.diagnostics.lastModel = "";
  state.diagnostics.lastProvider = "";
  state.diagnostics.lastLatencyMs = 0;
}

export function getDiagnosticsSnapshot() {
  return JSON.parse(JSON.stringify(state.diagnostics || {}));
}

export function formatDiagnosticsForUi(snapshot) {
  // Gardé pour compatibilité, mais l'UI utilise maintenant getDiagnosticsSnapshot directement
  if (!snapshot) return "";
  return `Requests: ${snapshot.requests} | Cost: $${(snapshot.estimatedCostUSD || 0).toFixed(4)}`;
}

export function getSharedState() {
  return state;
}
