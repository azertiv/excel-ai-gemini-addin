// src/taskpane/taskpane.js

import { getApiKey, setApiKey, clearApiKey, getMaxTokens, setMaxTokens, storageBackend } from "../shared/storage";
import { geminiMinimalTest } from "../shared/gemini";
import { getDiagnosticsSnapshot, resetDiagnosticsLogs } from "../shared/diagnostics";

let els = {};

function $(id) { return document.getElementById(id); }

// --- LOGIQUE ONGLETS ---
function initTabs() {
  document.querySelectorAll('.tab').forEach(t => {
    t.addEventListener('click', () => {
      // Switch active tab
      document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(x => x.classList.remove('active'));

      t.classList.add('active');
      const targetId = t.dataset.target;
      document.getElementById(targetId).classList.add('active');

      // Refresh specific tab data
      if(targetId === 'logs') updateLogsUI();
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
    if(el) {
        el.style.display = (el.style.display === 'block') ? 'none' : 'block';
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
  const padding = 24;
  const usableWidth = width - padding * 2;
  const usableHeight = height - padding * 2;

  ctx.strokeStyle = '#e5e5e5';
  ctx.lineWidth = 1;
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

  ctx.strokeStyle = '#0f6cbd';
  ctx.fillStyle = 'rgba(15, 108, 189, 0.12)';
  ctx.lineWidth = 2;
  ctx.beginPath();

  series.forEach((bucket, idx) => {
    const x = padding + (usableWidth / (series.length - 1)) * idx;
    const y = padding + (1 - bucket.value / maxVal) * usableHeight;
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

  ctx.fillStyle = '#555';
  ctx.font = '10px "Segoe UI", Arial, sans-serif';
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
      list.innerHTML = '<div style="padding:15px; text-align:center; color:#999; font-style:italic;">Aucune requête récente</div>';
      return;
    }

    let html = "";
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

            <div id="code-${uniqueId}" class="log-full-code">${safeDetailText}</div>
        </div>
      `;
    }
    list.innerHTML = html;

    // Attach event listeners for toggles
    list.querySelectorAll('.log-code-toggle').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = e.target.getAttribute('data-id');
            toggleLogDetails(id);
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
  const backend = await storageBackend();

  els.keyStatus.textContent = key ? "OK" : "MISSING";
  els.keyStatus.className = key ? "status ok" : "status missing";

  if (els.maxTokensInput) {
    els.maxTokensInput.value = maxTokens || "";
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
    const t = (els.maxTokensInput.value || "").trim();

    if (v) {
      await setApiKey(v);
      els.apiKeyInput.value = "";
    } else {
        const currentKey = await getApiKey();
        if (!currentKey) {
             setMessage("Collez une clé API Gemini valide.", "warn");
             return;
        }
    }

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
    const res = await geminiMinimalTest({ timeoutMs: 10000 });
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
  updateLogsUI();
}

function wireUi() {
  els = {
    apiKeyInput: $("apiKeyInput"),
    maxTokensInput: $("maxTokensInput"),
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
    usageChart: $("usageChart"),
    // Features UI
    featuresHeader: $("featuresHeader"),
    featuresContent: $("featuresContent"),
    featuresToggle: $("featuresToggle")
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

  els.apiKeyInput.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") onSave();
  });

  if (els.maxTokensInput) {
    els.maxTokensInput.addEventListener("keydown", (ev) => {
        if (ev.key === "Enter") onSave();
    });
  }

  // Features toggle
  if (els.featuresHeader) {
    els.featuresHeader.addEventListener("click", () => {
        const isHidden = els.featuresContent.style.display === "none";
        els.featuresContent.style.display = isHidden ? "block" : "none";
        els.featuresToggle.textContent = isHidden ? "▲" : "▼";
    });
  }

  initTabs();
}

Office.onReady(async () => {
  wireUi();
  await refreshKeyStatus();
  updateLogsUI();

  // Auto-refresh logs if tab is active
  setInterval(() => {
    const logsTab = document.getElementById('logs');
    if(logsTab && logsTab.classList.contains('active')) {
      updateLogsUI();
    }
  }, 4000);
});