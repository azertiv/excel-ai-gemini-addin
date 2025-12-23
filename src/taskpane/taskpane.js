// src/taskpane/taskpane.js

import { getApiKey, setApiKey, clearApiKey, getMaxTokens, setMaxTokens, storageBackend } from "../shared/storage";
import { geminiMinimalTest } from "../shared/gemini";
import { getDiagnosticsSnapshot } from "../shared/diagnostics";

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

function updateLogsUI() {
  try {
    const snap = getDiagnosticsSnapshot();

    // 1. Stats
    const total = (snap.totalInputTokens || 0) + (snap.totalOutputTokens || 0);
    const cost = (snap.estimatedCostUSD || 0);

    if (els.totalTokens) els.totalTokens.textContent = total.toLocaleString();
    if (els.estCost) els.estCost.textContent = '$' + cost.toFixed(5);

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
      const cacheBadge = log.cached ? '<span class="badge cached">CACHE</span>' : '';
      const tokensInfo = log.cached ? '-' : `${log.inputTokens} &rarr; ${log.outputTokens}`;

      html += `
        <div class="log-item ${isErr ? 'err' : ''}">
          <div class="log-time">${formatTime(log.at)}</div>

          <div class="log-main">
             ${cacheBadge}
             <span title="${log.model}">${log.model}</span>
             <span class="log-code">${log.code}</span>
          </div>

          <div class="log-meta">
             <div>${tokensInfo}</div>
             <div>${log.latencyMs}ms</div>
          </div>
        </div>
      `;
    }
    list.innerHTML = html;
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