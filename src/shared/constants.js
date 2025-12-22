export const STORAGE = {
  API_KEY: "AI_GEMINI_API_KEY_V1",
  PERSIST_CACHE_INDEX: "AI_PERSIST_CACHE_INDEX_V1"
};

export const GEMINI = {
  BASE_URL: "https://generativelanguage.googleapis.com/v1beta",
  DEFAULT_MODEL: "gemini-3-flash-preview"
};

export const LIMITS = {
  MAX_CONTEXT_CHARS: 3500,
  MAX_INPUT_CHARS: 12000,
  MAX_CELL_CHARS: 32000,

  MAX_TABLE_ROWS: 50,
  MAX_TABLE_COLS: 12,

  MAX_FILL_ROWS: 200,
  MAX_EXAMPLES: 40,

  MEM_CACHE_ENTRIES: 200,
  MEM_CACHE_TTL_MS: 60 * 60 * 1000, // 1h

  MAX_CONCURRENT_REQUESTS: 3
};

export const DEFAULTS = {
  lang: "fr",
  timeoutMs: 20000,
  retry: 1,
  cache: "memory",
  cacheTtlSec: 3600,
  temperature: 0.2,
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
  TOO_LARGE: "#AI_TOO_LARGE",
  EMPTY_RESPONSE: "#AI_EMPTY_RESPONSE"
};
