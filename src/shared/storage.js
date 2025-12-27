import { STORAGE, TOKEN_LIMITS, PROVIDERS, DEFAULTS } from "./constants";
import { diagSet } from "./diagnostics";

let _officeReadyPromise = null;

function officeReady() {
  if (_officeReadyPromise) return _officeReadyPromise;
  try {
    if (typeof Office !== "undefined" && Office.onReady) _officeReadyPromise = Office.onReady();
    else _officeReadyPromise = Promise.resolve();
  } catch {
    _officeReadyPromise = Promise.resolve();
  }
  return _officeReadyPromise;
}

async function detectBackend() {
  await officeReady();
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && typeof OfficeRuntime.storage.getItem === "function") {
      diagSet("backend", "OfficeRuntime.storage");
      return "office";
    }
  } catch { /* ignore */ }

  try {
    if (typeof localStorage !== "undefined") {
      diagSet("backend", "localStorage");
      return "local";
    }
  } catch { /* ignore */ }

  diagSet("backend", "none");
  return "none";
}

let _backendPromise = null;
function backend() {
  if (!_backendPromise) _backendPromise = detectBackend();
  return _backendPromise;
}

export async function storageBackend() {
  return await backend();
}

export async function getItem(key) {
  const b = await backend();
  if (b === "office") return await OfficeRuntime.storage.getItem(key);
  if (b === "local") return localStorage.getItem(key);
  return null;
}

export async function setItem(key, value) {
  const b = await backend();
  if (b === "office") { await OfficeRuntime.storage.setItem(key, value); return; }
  if (b === "local") localStorage.setItem(key, value);
}

export async function removeItem(key) {
  const b = await backend();
  if (b === "office") { await OfficeRuntime.storage.removeItem(key); return; }
  if (b === "local") localStorage.removeItem(key);
}

const API_KEYS = {
  [PROVIDERS.GEMINI]: {
    storageKey: STORAGE.API_KEY,
    state: { loaded: false, value: "", promise: null }
  },
  [PROVIDERS.OPENAI]: {
    storageKey: STORAGE.API_KEY_OPENAI,
    state: { loaded: false, value: "", promise: null }
  }
};

function normProvider(p) {
  return p === PROVIDERS.OPENAI ? PROVIDERS.OPENAI : PROVIDERS.GEMINI;
}

async function readString(storageKey) {
  const v = await getItem(storageKey);
  return typeof v === "string" ? v : String(v || "");
}

export async function getApiKey(provider = PROVIDERS.GEMINI) {
  const p = normProvider(provider);
  const { storageKey, state } = API_KEYS[p];

  if (state.loaded) return state.value;
  if (state.promise) return await state.promise;

  state.promise = (async () => {
    const val = await readString(storageKey);
    state.value = val.trim();
    state.loaded = true;
    state.promise = null;
    return state.value;
  })();

  return await state.promise;
}

export async function setApiKey(provider, apiKey) {
  const p = normProvider(provider);
  const { storageKey, state } = API_KEYS[p];
  const key = (apiKey || "").trim();

  state.value = key;
  state.loaded = true;
  state.promise = null;

  if (!key) await removeItem(storageKey);
  else await setItem(storageKey, key);
  return true;
}

export async function clearApiKey(provider = PROVIDERS.GEMINI) {
  const p = normProvider(provider);
  const { storageKey, state } = API_KEYS[p];

  state.value = "";
  state.loaded = true;
  state.promise = null;

  await removeItem(storageKey);
  return true;
}

let _maxTokensLoaded = false;
let _maxTokensValue = null;
let _maxTokensLoadPromise = null;

export async function getMaxTokens() {
  if (_maxTokensLoaded) return _maxTokensValue;
  if (_maxTokensLoadPromise) return await _maxTokensLoadPromise;

  _maxTokensLoadPromise = (async () => {
    const v = await getItem(STORAGE.MAX_TOKENS);
    const n = Math.floor(Number(v));
    if (Number.isFinite(n) && n >= TOKEN_LIMITS.MIN) {
      _maxTokensValue = Math.min(TOKEN_LIMITS.MAX, Math.max(TOKEN_LIMITS.MIN, n));
    } else {
      _maxTokensValue = null;
    }
    _maxTokensLoaded = true;
    return _maxTokensValue;
  })();

  return await _maxTokensLoadPromise;
}

export async function setMaxTokens(val) {
  const n = Math.floor(Number(val));
  const valid = Number.isFinite(n) && n >= TOKEN_LIMITS.MIN;

  if (valid) {
    const clamped = Math.min(TOKEN_LIMITS.MAX, Math.max(TOKEN_LIMITS.MIN, n));
    _maxTokensValue = clamped;
    await setItem(STORAGE.MAX_TOKENS, String(clamped));
  } else {
    _maxTokensValue = null;
    await removeItem(STORAGE.MAX_TOKENS);
  }

  _maxTokensLoaded = true;
  _maxTokensLoadPromise = null;
  return true;
}

// Provider preference -------------------------------------------------
let _providerLoaded = false;
let _providerValue = DEFAULTS.provider;
let _providerLoadPromise = null;

export async function getProvider() {
  if (_providerLoaded) return _providerValue;
  if (_providerLoadPromise) return await _providerLoadPromise;

  _providerLoadPromise = (async () => {
    const v = await getItem(STORAGE.PROVIDER);
    const val = typeof v === "string" ? v.trim().toLowerCase() : "";
    _providerValue = val === PROVIDERS.OPENAI ? PROVIDERS.OPENAI : PROVIDERS.GEMINI;
    _providerLoaded = true;
    _providerLoadPromise = null;
    return _providerValue;
  })();

  return await _providerLoadPromise;
}

export async function setProvider(provider) {
  const p = normProvider(provider);
  _providerValue = p;
  _providerLoaded = true;
  _providerLoadPromise = null;
  await setItem(STORAGE.PROVIDER, p);
  return true;
}

// Model preferences ---------------------------------------------------
const MODEL_KEYS = {
  [PROVIDERS.GEMINI]: STORAGE.GEMINI_MODEL,
  [PROVIDERS.OPENAI]: STORAGE.OPENAI_MODEL
};
const modelState = {
  [PROVIDERS.GEMINI]: { loaded: false, value: "", promise: null },
  [PROVIDERS.OPENAI]: { loaded: false, value: "", promise: null }
};

export async function getModel(provider = PROVIDERS.GEMINI) {
  const p = normProvider(provider);
  const state = modelState[p];
  if (state.loaded) return state.value;
  if (state.promise) return await state.promise;

  state.promise = (async () => {
    const v = await getItem(MODEL_KEYS[p]);
    state.value = typeof v === "string" ? v.trim() : "";
    state.loaded = true;
    state.promise = null;
    return state.value;
  })();
  return await state.promise;
}

export async function setModel(provider, model) {
  const p = normProvider(provider);
  const state = modelState[p];
  const key = MODEL_KEYS[p];
  const m = (model || "").trim();

  state.value = m;
  state.loaded = true;
  state.promise = null;

  if (m) await setItem(key, m);
  else await removeItem(key);
  return true;
}
