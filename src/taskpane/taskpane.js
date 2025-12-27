// src/taskpane/taskpane.js

import { getApiKey, setApiKey, clearApiKey, getMaxTokens, setMaxTokens, getBaseUrl, setBaseUrl, storageBackend } from "../shared/storage";
import { openaiMinimalTest } from "../shared/openai";
import { getDiagnosticsSnapshot, resetDiagnosticsLogs } from "../shared/diagnostics";
import { DEFAULTS, TOKEN_LIMITS } from "../shared/constants";

const TOKEN_STEPS = (() => {
  const steps = [];
  let value = 64;
  while (value < TOKEN_LIMITS.MAX) {
    steps.push(value);
    value *= 2;
  }
  if (!steps.includes(TOKEN_LIMITS.MAX)) steps.push(TOKEN_LIMITS.MAX);
  return steps;
})();

let els = {};
let openLogDetails = new Set();

function $(id) { return document.getElementById(id); }

function toggleSectionVisibility(targetId, collapsed) {
  const content = document.getElementById(targetId);
  if (!content) return;

  const shouldCollapse = typeof collapsed === "boolean"
    ? collapsed
    : content.style.display !== "none";

  content.style.display = shouldCollapse ? "none" : "block";
  content.classList.toggle("collapsed", shouldCollapse);

  const toggles = document.querySelectorAll(`[data-toggle-section="${targetId}"]`);
  toggles.forEach(btn => {
    const chevron = btn.querySelector('.chevron');
    const label = btn.querySelector('.collapse-label');
    if (chevron) chevron.textContent = shouldCollapse ? '▼' : '▲';
    if (label) label.textContent = shouldCollapse ? 'Déplier' : 'Réduire';
  });
}

function clampTokenValue(raw) {
  const n = Math.floor(Number(raw));
  if (!Number.isFinite(n)) return DEFAULTS.maxTokens;
  let closest = TOKEN_STEPS[0];
  for (const step of TOKEN_STEPS) {
    if (Math.abs(step - n) < Math.abs(closest - n)) {
      closest = step;
    }
  }
  return Math.max(TOKEN_LIMITS.MIN, Math.min(TOKEN_LIMITS.MAX, closest));
}

function sliderIndexForValue(value) {
  const v = clampTokenValue(value);
  const idx = TOKEN_STEPS.indexOf(v);
  return idx >= 0 ? idx : 0;
}

function sliderValueToTokens(sliderVal) {
  const idx = Math.max(0, Math.min(TOKEN_STEPS.length - 1, Number(sliderVal) || 0));
  return TOKEN_STEPS[idx];
}

function setTokenUIValue(raw) {
  const v = clampTokenValue(raw);
  if (els.maxTokensSlider) els.maxTokensSlider.value = String(sliderIndexForValue(v));
  if (els.maxTokensInput) els.maxTokensInput.value = String(v);
  if (els.maxTokensValue) els.maxTokensValue.textContent = String(v);
  return v;
}

function setupSectionToggles() {
  document.querySelectorAll('[data-toggle-section]').forEach(btn => {
    const targetId = btn.dataset.toggleSection;
    const content = document.getElementById(targetId);
    if (!targetId || !content) return;

    const startCollapsed = content.dataset.startCollapsed === "true";
    if (!content.dataset.collapseInit) {
      toggleSectionVisibility(targetId, startCollapsed);
      content.dataset.collapseInit = "1";
    }

    btn.addEventListener('click', () => {
      const currentlyHidden = content.style.display === 'none';
      toggleSectionVisibility(targetId, !currentlyHidden);
    });
  });
}

// --- LOGIQUE LOGS ---
function formatTime(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  return d.toLocaleTimeString([], { hour: '2-digit', minute:'2-digit', second:'2-digit' });
}

function formatLatencySeconds(latencyMs) {
  const seconds = (latencyMs || 0) / 1000;
  return `${seconds.toFixed(2)}s`;
}

function getTokenColorClass(total) {
  if (total < 500) return 'low';
  if (total < 2000) return 'medium';
  return 'high';
}

// Helper to toggle log details
function toggleLogDetails(id) {
  const el = document.getElementById('code-' + id);
  if (!el) return;

  const isOpen = el.style.display === 'block';
  const nextState = !isOpen;
  el.style.display = nextState ? 'block' : 'none';

  if (nextState) {
    openLogDetails.add(id);
  } else {
    openLogDetails.delete(id);
  }
}

// Ensure function is safe to call
function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function computeUsageSeries(logs) {
  const bucketCount = 12; // 5 minutes per bucket over the last hour
  const now = Date.now();
  const hourAgo = now - 60 * 60 * 1000;
  const bucketMs = (60 * 60 * 1000) / bucketCount;
  const buckets = Array.from({ length: bucketCount }, (_, idx) => ({
    ts: hourAgo + idx * bucketMs,
    value: 0
  }));

  logs
    .filter(l => new Date(l.at).getTime() >= hourAgo)
    .forEach(log => {
      const ts = new Date(log.at).getTime();
      const bucketIndex = Math.min(bucketCount - 1, Math.floor((ts - hourAgo) / bucketMs));
      const totalTokens = (log.inputTokens || 0) + (log.outputTokens || 0);
      buckets[bucketIndex].value += totalTokens;
    });

  return buckets;
}

function renderUsageChart(logs) {
  if (!els.usageChart) return;

  const ctx = els.usageChart.getContext('2d');
  if (!ctx) return;

  const width = els.usageChart.clientWidth || 360;
  const height = els.usageChart.clientHeight || 160;
  els.usageChart.width = width;
  els.usageChart.height = height;

  ctx.clearRect(0, 0, width, height);

  const series = computeUsageSeries(logs || []);
  const maxVal = Math.max(...series.map(s => s.value), 1);
  const niceMax = (() => {
    const magnitude = Math.pow(10, Math.floor(Math.log10(maxVal)));
    const scaled = maxVal / magnitude;
    if (scaled <= 1.5) return 2 * magnitude;
    if (scaled <= 3) return 4 * magnitude;
    if (scaled <= 7.5) return 8 * magnitude;
    return 10 * magnitude;
  })();

  const padding = 24;
  const usableWidth = width - padding * 2;
  const usableHeight = height - padding * 2;

  // Grid
  ctx.strokeStyle = '#e9eef5';
  ctx.lineWidth = 1;
  const gridLines = 4;
  for (let i = 0; i <= gridLines; i++) {
    const y = padding + (usableHeight / gridLines) * i;
    ctx.beginPath();
    ctx.moveTo(padding, y);
    ctx.lineTo(width - padding, y);
    ctx.stroke();
  }

  // X axis baseline
  ctx.strokeStyle = '#ccd8ea';
  ctx.beginPath();
  ctx.moveTo(padding, height - padding);
  ctx.lineTo(width - padding, height - padding);
  ctx.stroke();

  if (series.every(s => s.value === 0)) {
    ctx.fillStyle = '#999';
    ctx.font = '12px "Segoe UI", Arial, sans-serif';
    ctx.fillText('Aucune activité sur la dernière heure', padding, height / 2);
    return;
  }

  const gradient = ctx.createLinearGradient(0, padding, 0, height - padding);
  gradient.addColorStop(0, 'rgba(15, 108, 189, 0.16)');
  gradient.addColorStop(1, 'rgba(15, 108, 189, 0.02)');

  ctx.strokeStyle = '#0f6cbd';
  ctx.fillStyle = gradient;
  ctx.lineWidth = 2;
  ctx.beginPath();

  series.forEach((bucket, idx) => {
    const x = padding + (usableWidth / (series.length - 1)) * idx;
    const y = padding + (1 - bucket.value / niceMax) * usableHeight;
    if (idx === 0) {
      ctx.moveTo(x, y);
    } else {
      ctx.lineTo(x, y);
    }
  });

  ctx.stroke();

  ctx.lineTo(width - padding, height - padding);
  ctx.lineTo(padding, height - padding);
  ctx.closePath();
  ctx.fill();

  // Points
  ctx.fillStyle = '#0f6cbd';
  series.forEach((bucket, idx) => {
    const x = padding + (usableWidth / (series.length - 1)) * idx;
    const y = padding + (1 - bucket.value / niceMax) * usableHeight;
    ctx.beginPath();
    ctx.arc(x, y, 3, 0, Math.PI * 2);
    ctx.fill();
  });

  // Axis labels
  ctx.fillStyle = '#555';
  ctx.font = '10px "Segoe UI", Arial, sans-serif';
  const labels = [
    { text: '-60m', x: padding },
    { text: '-30m', x: padding + usableWidth / 2 },
    { text: 'Maintenant', x: width - padding }
  ];
  labels.forEach((l, idx) => {
    ctx.textAlign = idx === 0 ? 'left' : idx === labels.length - 1 ? 'right' : 'center';
    ctx.fillText(l.text, l.x, height - padding + 14);
  });

  ctx.textAlign = 'left';
  ctx.fillText('Tokens (dernière heure)', padding, padding - 6);
}

function updateLogsUI() {
  try {
    const snap = getDiagnosticsSnapshot();

    // 1. Stats
    const total = (snap.totalInputTokens || 0) + (snap.totalOutputTokens || 0);
    const cost = (snap.estimatedCostUSD || 0);

    if (els.totalTokens) els.totalTokens.textContent = total.toLocaleString();
    if (els.estCost) els.estCost.textContent = '$' + cost.toFixed(5);

    renderUsageChart(snap.logs || []);

    // 2. Liste Logs
    const list = els.logList;
    if (!list) return;

    if (!snap.logs || snap.logs.length === 0) {
      openLogDetails.clear();
      list.innerHTML = '<div style="padding:15px; text-align:center; color:#999; font-style:italic;">Aucune requête récente</div>';
      return;
    }

    let html = "";
    const activeIds = new Set();
    for (const log of snap.logs) {
      const isErr = !log.success;
      const logTotalTokens = (log.inputTokens || 0) + (log.outputTokens || 0);
      const tokenClass = getTokenColorClass(logTotalTokens);
      const cacheBadge = log.cached ? '<span class="badge cached">CACHE</span>' : '';
      const errBadge = isErr ? '<span class="badge err">ERR</span>' : '';

      const funcName = log.functionName || log.code || "UNKNOWN";
      // Si log.code est different de funcName et de "OK", on peut vouloir l'afficher (ex: type d'erreur)
      const errCode = (isErr && log.code !== funcName) ? `(${log.code})` : "";

      const uniqueId = log.id || Math.random().toString(36).substr(2, 9);
      const isOpen = openLogDetails.has(uniqueId);
      activeIds.add(uniqueId);

      // Full details (Formula/Code/Prompt)
      const detailText = log.message || "(No details available)";
      // Security: escape content
      const safeDetailText = escapeHtml(detailText);

      html += `
        <div class="log-item ${isErr ? 'err' : ''}">
            <div class="log-row-top">
                <!-- Left: Token Badge -->
                <div class="token-badge ${tokenClass}">
                    <div>${log.cached ? 'CACHE' : logTotalTokens}</div>
                    ${!log.cached ? `<div class="token-details">${log.inputTokens} &rarr; ${log.outputTokens}</div>` : ''}
                </div>

                <!-- Middle: Main Info -->
                <div class="log-info">
                    <div>
                        <span class="log-func">${funcName}</span>
                        <span class="log-time">${formatTime(log.at)}</span>
                    </div>
                    <div class="log-model">
                        ${log.model} ${errCode}
                    </div>
                    ${detailText ? `<div class="log-code-toggle" data-id="${uniqueId}">Afficher détails</div>` : ''}
                </div>

                <!-- Right: Status -->
                <div class="log-status">
                    <div class="log-latency">${formatLatencySeconds(log.latencyMs)}</div>
                    ${cacheBadge}
                    ${errBadge}
                </div>
            </div>

            <div id="code-${uniqueId}" class="log-full-code" style="display:${isOpen ? 'block' : 'none'}">
              <div class="log-full-code__actions">
                <button class="copy-btn" data-copy-id="${uniqueId}" title="Copier le prompt">📋</button>
              </div>
              <pre class="log-full-code__text">${safeDetailText}</pre>
            </div>
        </div>
      `;
    }
    list.innerHTML = html;

    // Nettoie les IDs qui ne sont plus présents (évite de rouvrir après reset)
    openLogDetails = new Set([...openLogDetails].filter(id => activeIds.has(id)));

    // Attach event listeners for toggles
    list.querySelectorAll('.log-code-toggle').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = e.target.getAttribute('data-id');
        toggleLogDetails(id);
      });
    });

    list.querySelectorAll('.copy-btn').forEach(btn => {
      btn.addEventListener('click', async (e) => {
        const id = e.currentTarget.getAttribute('data-copy-id');
        const textEl = document.getElementById('code-' + id)?.querySelector('.log-full-code__text');
        if (!textEl) return;
        try {
          await navigator.clipboard.writeText(textEl.textContent || '');
          setMessage('Prompt copié dans le presse-papiers.', 'info');
        } catch (err) {
          console.error('Clipboard error', err);
          setMessage('Impossible de copier le prompt.', 'error');
        }
      });
    });
  } catch (e) {
    console.error("Error updating logs UI", e);
  }
}

// --- LOGIQUE SETTINGS ---
function setTestDiagnostics(text) {
  if (!els.testDiag) return;
  if (text) {
    els.testDiag.textContent = text;
    els.testDiag.style.display = "block";
  } else {
    els.testDiag.textContent = "";
    els.testDiag.style.display = "none";
  }
}

function formatTestDiagnostics(res) {
  const d = res?.diagnostics || {};
  const lines = [];

  if (res?.message) lines.push(`message: ${res.message}`);
  if (typeof d.httpStatus === "number") lines.push(`httpStatus: ${d.httpStatus}`);
  if (typeof d.candidates === "number") lines.push(`candidates: ${d.candidates}`);
  if (d.finishReason) lines.push(`finishReason: ${d.finishReason}`);
  if (d.cacheSource) lines.push(`cacheSource: ${d.cacheSource}`);
  if (res?.cacheKey) lines.push(`cacheKey: ${res.cacheKey}`);
  if (d.latencyMs) lines.push(`latencyMs: ${d.latencyMs}`);

  return lines.join("\n") || "(no diagnostics)";
}

async function refreshKeyStatus() {
  const key = (await getApiKey()) || "";
  const maxTokens = await getMaxTokens();
  const baseUrl = await getBaseUrl();
  const backend = await storageBackend();

  els.keyStatus.textContent = key ? "OK" : "MISSING";
  els.keyStatus.className = key ? "status ok" : "status missing";

  setTokenUIValue(maxTokens ?? DEFAULTS.maxTokens);

  if (els.baseUrlInput) {
    els.baseUrlInput.value = baseUrl === "https://api.openai.com/v1" ? "" : baseUrl;
  }

  els.backend.textContent =
    backend === "office" ? "OfficeRuntime.storage" :
    backend === "local" ? "localStorage" :
    "none";
}

function setMessage(msg, kind = "info") {
  els.message.textContent = msg || "";
  els.message.className = kind ? `message ${kind}` : "message";
}

async function onSave() {
  try {
    const v = (els.apiKeyInput.value || "").trim();
    const baseUrl = (els.baseUrlInput.value || "").trim();
    const sliderTokens = els.maxTokensSlider ? sliderValueToTokens(els.maxTokensSlider.value) : undefined;
    const rawTokenInput = (els.maxTokensInput?.value || "").trim();
    const preferredValue = rawTokenInput !== "" ? rawTokenInput : sliderTokens;
    const t = setTokenUIValue(preferredValue ?? DEFAULTS.maxTokens);

    if (v) {
      await setApiKey(v);
      els.apiKeyInput.value = "";
    } else {
        const currentKey = await getApiKey();
        if (!currentKey) {
             setMessage("Collez une clé API OpenAI valide.", "warn");
             return;
        }
    }

    await setBaseUrl(baseUrl);
    await setMaxTokens(t);

    setMessage("Configuration sauvegardée.", "ok");
    await refreshKeyStatus();
  } catch (e) {
    setMessage("Erreur sauvegarde.", "error");
    console.error(e);
  }
}

async function onClear() {
  try {
    await clearApiKey();
    els.apiKeyInput.value = "";
    setMessage("Clé supprimée.", "ok");
    await refreshKeyStatus();
  } catch (e) {
    console.error(e);
  }
}

async function onTest() {
  setMessage("Test en cours...", "info");
  setTestDiagnostics("");
  try {
    const res = await openaiMinimalTest({ timeoutMs: 10000 });
    if (res.ok) {
      setMessage("Test API : OK", "ok");
      setTestDiagnostics(formatTestDiagnostics(res));
    } else {
      setMessage(`Erreur : ${res.code}`, "error");
      setTestDiagnostics(formatTestDiagnostics(res));
    }
    // Update logs after test
    updateLogsUI();
  } catch (e) {
    setMessage("Erreur inattendue", "error");
    setTestDiagnostics(e?.message);
  } finally {
    await refreshKeyStatus();
  }
}

  function onResetLogs() {
    const confirmed = window.confirm('Réinitialiser les logs et les coûts ?');
    if (!confirmed) return;

    resetDiagnosticsLogs();
    openLogDetails.clear();
    updateLogsUI();
    setMessage('Historique et compteurs remis à zéro.', 'info');
  }

function wireUi() {
  els = {
    apiKeyInput: $("apiKeyInput"),
    baseUrlInput: $("baseUrlInput"),
    maxTokensInput: $("maxTokensInput"),
    maxTokensSlider: $("maxTokensSlider"),
    maxTokensValue: $("maxTokensValue"),
    saveBtn: $("saveKeyBtn"),
    clearBtn: $("clearKeyBtn"),
    testBtn: $("testBtn"),
    keyStatus: $("keyStatus"),
    backend: $("backend"),
    message: $("message"),
    testDiag: $("testDiagnostics"),
    // Logs UI
    totalTokens: $("totalTokens"),
    estCost: $("estCost"),
    logList: $("logList"),
    refreshLogsBtn: $("refreshLogsBtn"),
    resetLogsBtn: $("resetLogsBtn"),
    usageChart: $("usageChart")
  };

  els.saveBtn.addEventListener("click", onSave);
  els.clearBtn.addEventListener("click", onClear);
  els.testBtn.addEventListener("click", onTest);

  if(els.refreshLogsBtn) {
    els.refreshLogsBtn.addEventListener("click", updateLogsUI);
  }

  if (els.resetLogsBtn) {
    els.resetLogsBtn.addEventListener("click", onResetLogs);
  }

  setupSectionToggles();

  els.apiKeyInput.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") onSave();
  });

  if (els.maxTokensInput) {
    els.maxTokensInput.addEventListener("keydown", (ev) => {
        if (ev.key === "Enter") onSave();
    });

    // Keep slider + numeric input synchronized.
    els.maxTokensInput.addEventListener("input", () => {
      setTokenUIValue(els.maxTokensInput.value);
    });
  }

  if (els.maxTokensSlider) {
    els.maxTokensSlider.setAttribute('min', '0');
    els.maxTokensSlider.setAttribute('max', String(Math.max(0, TOKEN_STEPS.length - 1)));
    els.maxTokensSlider.setAttribute('step', '1');
    els.maxTokensSlider.addEventListener("input", () => {
      const snapped = sliderValueToTokens(els.maxTokensSlider.value);
      setTokenUIValue(snapped);
    });
  }

}

Office.onReady(async () => {
  wireUi();
  await refreshKeyStatus();
  updateLogsUI();

  // Auto-refresh logs
  setInterval(() => {
    updateLogsUI();
  }, 4000);
});