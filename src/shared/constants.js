export const STORAGE = {
  API_KEY: "AI_OPENAI_API_KEY_V1",
  MAX_TOKENS: "AI_OPENAI_MAX_TOKENS_V1",
  API_BASE_URL: "AI_OPENAI_API_BASE_URL_V1",
  PREFERRED_MODEL: "AI_OPENAI_MODEL_V1",
  PERSIST_CACHE_INDEX: "AI_PERSIST_CACHE_INDEX_V1"
};

export const OPENAI = {
  BASE_URL: "https://api.openai.com/v1",
  DEFAULT_MODEL: "gpt-5.0-mini"
};

// Global output token limit (maxOutputTokens) bounds exposed in the taskpane.
// NOTE: This controls the model OUTPUT tokens. Input/context is only limited by the model context window.
export const TOKEN_LIMITS = {
  MIN: 32,
  MAX: 128000,
  STEP: 32
};

export const LIMITS = {
  // Excel hard limit: a cell can contain up to 32,767 characters.
  // We enforce this only when returning a single-cell text result.
  MAX_CELL_CHARS: 32767,

  // Cache / runtime safety (not an output/content limitation)
  MEM_CACHE_ENTRIES: 200,
  MEM_CACHE_TTL_MS: 60 * 60 * 1000, // 1h

  // Avoid flooding the model API with too many concurrent requests
  MAX_CONCURRENT_REQUESTS: 3
};

export const DEFAULTS = {
  lang: "fr",
  timeoutMs: 120000,
  retry: 1,
  cache: "persistent",
  cacheTtlSec: 24 * 3600,
  temperature: 0.2,
  // Used only when no stored setting is present and no per-formula option is provided.
  maxTokens: 2048
};

export const ERR = {
  KEY_MISSING: "#AI_KEY_MISSING",
  BAD_INPUT: "#AI_BAD_INPUT",
  BAD_OPTIONS: "#AI_BAD_OPTIONS",
  BAD_SCHEMA: "#AI_BAD_SCHEMA",
  TIMEOUT: "#AI_TIMEOUT",
  RATE_LIMIT: "#AI_RATE_LIMIT",
  AUTH: "#AI_AUTH",
  BLOCKED: "#AI_BLOCKED",
  API_ERROR: "#AI_API_ERROR",
  PARSE_ERROR: "#AI_PARSE_ERROR",
  NOT_FOUND: "#AI_NOT_FOUND",
  CACHE_MISS: "#AI_CACHE_MISS",
  TOO_LARGE: "#AI_TOO_LARGE",
  EMPTY_RESPONSE: "#AI_EMPTY_RESPONSE"
};
