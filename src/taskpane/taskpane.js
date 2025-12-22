import { getApiKey, setApiKey, clearApiKey, storageBackend } from "../shared/storage";
import { geminiMinimalTest } from "../shared/gemini";
import { getDiagnosticsSnapshot, formatDiagnosticsForUi } from "../shared/diagnostics";

let els = {};

function $(id) { return document.getElementById(id); }

async function refreshKeyStatus() {
  const key = (await getApiKey()) || "";
  const backend = await storageBackend();

  els.keyStatus.textContent = key ? "OK" : "MISSING";
  els.keyStatus.className = key ? "status ok" : "status missing";

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
    if (!v) {
      setMessage("Collez une clé API Gemini valide, puis cliquez sur Save.", "warn");
      return;
    }
    await setApiKey(v);
    els.apiKeyInput.value = "";
    setMessage("Clé API sauvegardée localement (masquée).", "ok");
    await refreshKeyStatus();
  } catch (e) {
    setMessage("Impossible de sauvegarder la clé. Voir console.", "error");
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
    setMessage("Impossible de supprimer la clé. Voir console.", "error");
    console.error(e);
  }
}

async function onTest() {
  setMessage("Test API en cours…", "info");
  try {
    const res = await geminiMinimalTest({ timeoutMs: 10000 });
    if (res.ok) {
      setMessage("Test API : OK", "ok");
    } else {
      setMessage(`Test API : ${res.code} (${res.message || "erreur"})`, "error");
    }
  } catch (e) {
    setMessage("Test API : erreur inattendue. Voir console.", "error");
    console.error(e);
  } finally {
    await refreshKeyStatus();
  }
}

function refreshDiagnostics() {
  try {
    const snap = getDiagnosticsSnapshot();
    els.diag.textContent = formatDiagnosticsForUi(snap);
  } catch {
    // ignore
  }
}

function wireUi() {
  els = {
    apiKeyInput: $("apiKeyInput"),
    saveBtn: $("saveKeyBtn"),
    clearBtn: $("clearKeyBtn"),
    testBtn: $("testBtn"),
    keyStatus: $("keyStatus"),
    backend: $("backend"),
    message: $("message"),
    diag: $("diag")
  };

  els.saveBtn.addEventListener("click", onSave);
  els.clearBtn.addEventListener("click", onClear);
  els.testBtn.addEventListener("click", onTest);

  els.apiKeyInput.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") onSave();
  });
}

Office.onReady(async () => {
  wireUi();
  await refreshKeyStatus();
  refreshDiagnostics();
  setInterval(refreshDiagnostics, 1500);
});
