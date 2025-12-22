export class LRUCache {
  constructor(maxEntries = 200, ttlMs = 0) {
    this.maxEntries = maxEntries;
    this.ttlMs = ttlMs;
    this.map = new Map();
  }

  get(key) {
    const entry = this.map.get(key);
    if (!entry) return undefined;

    if (this.ttlMs > 0 && Date.now() - entry.t > this.ttlMs) {
      this.map.delete(key);
      return undefined;
    }

    this.map.delete(key);
    this.map.set(key, entry);
    return entry.v;
  }

  set(key, value) {
    if (this.map.has(key)) this.map.delete(key);
    this.map.set(key, { v: value, t: Date.now() });

    while (this.map.size > this.maxEntries) {
      const oldestKey = this.map.keys().next().value;
      this.map.delete(oldestKey);
    }
  }

  clear() { this.map.clear(); }

  stats() {
    return { size: this.map.size, maxEntries: this.maxEntries, ttlMs: this.ttlMs };
  }
}
