import { STORAGE, TOKEN_LIMITS } from "./constants";
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

let _apiKeyLoaded = false;
let _apiKeyValue = "";
let _apiKeyLoadPromise = null;

export async function getApiKey() {
  if (_apiKeyLoaded) return _apiKeyValue;
  if (_apiKeyLoadPromise) return await _apiKeyLoadPromise;

  _apiKeyLoadPromise = (async () => {
    const v = (await getItem(STORAGE.API_KEY)) || "";
    _apiKeyValue = typeof v === "string" ? v : String(v || "");
    _apiKeyLoaded = true;
    return _apiKeyValue;
  })();

  return await _apiKeyLoadPromise;
}

export async function setApiKey(apiKey) {
  const key = (apiKey || "").trim();
  _apiKeyValue = key;
  _apiKeyLoaded = true;
  _apiKeyLoadPromise = null;

  if (!key) await removeItem(STORAGE.API_KEY);
  else await setItem(STORAGE.API_KEY, key);

  return true;
}

export async function clearApiKey() {
  _apiKeyValue = "";
  _apiKeyLoaded = true;
  _apiKeyLoadPromise = null;
  await removeItem(STORAGE.API_KEY);
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
