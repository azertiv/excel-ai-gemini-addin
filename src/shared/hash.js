function fnv1a32(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return ("00000000" + h.toString(16)).slice(-8);
}

async function sha256Hex(str) {
  try {
    if (typeof crypto === "undefined" || !crypto.subtle || typeof TextEncoder === "undefined") return null;
    const data = new TextEncoder().encode(str);
    const hash = await crypto.subtle.digest("SHA-256", data);
    const bytes = new Uint8Array(hash);
    let hex = "";
    for (const b of bytes) hex += ("00" + b.toString(16)).slice(-2);
    return hex;
  } catch {
    return null;
  }
}

export async function hashKey(str) {
  const sha = await sha256Hex(str);
  return sha || fnv1a32(str);
}
