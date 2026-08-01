/* PDF Quick Disassembler — Web UI JS  GoTopo 2026 */

"use strict";

// ── Tab switching ─────────────────────────────────────────────────────────────
document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
  });
});

// ── Drop zone setup ───────────────────────────────────────────────────────────
function setupDropZone(zoneId, fileInputId, nameId, runBtnId) {
  const zone     = document.getElementById(zoneId);
  const input    = document.getElementById(fileInputId);
  const nameEl   = document.getElementById(nameId);
  const runBtn   = document.getElementById(runBtnId);
  let   selected = null;

  function setFile(file) {
    selected = file;
    nameEl.textContent = "✅  " + file.name;
    zone.classList.add("has-file");
    runBtn.disabled = false;
  }

  // File input change
  input.addEventListener("change", () => {
    if (input.files[0]) setFile(input.files[0]);
  });

  // Drag & drop
  zone.addEventListener("dragover", e => {
    e.preventDefault();
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });

  return () => selected;
}

const getFile1 = setupDropZone("drop1", "file1", "name1", "run1");
const getFile2 = setupDropZone("drop2", "file2", "name2", "run2");
const getFile3 = setupDropZone("drop3", "file3", "name3", "run3");

// ── Progress polling ──────────────────────────────────────────────────────────
function pollJob(jobId, fillId, msgId, onDone) {
  const fill = document.getElementById(fillId);
  const msg  = document.getElementById(msgId);

  const interval = setInterval(async () => {
    try {
      const res  = await fetch("/api/job/" + jobId);
      if (!res.ok) { clearInterval(interval); return; }
      const data = await res.json();

      fill.style.width = (data.pct || 0) + "%";
      msg.textContent  = data.msg || "";

      if (data.status === "done" || data.status === "error") {
        clearInterval(interval);
        onDone(data);
      }
    } catch (e) {
      clearInterval(interval);
    }
  }, 600);
}

// ── Generic run handler ───────────────────────────────────────────────────────
function setupRun(btnId, progWrapId, fillId, msgId, resultId, getFile, buildForm, endpoint) {
  const btn      = document.getElementById(btnId);
  const progWrap = document.getElementById(progWrapId);
  const fill     = document.getElementById(fillId);
  const resultEl = document.getElementById(resultId);

  btn.addEventListener("click", async () => {
    const file = getFile();
    if (!file) return;

    btn.disabled     = true;
    progWrap.style.display = "block";
    resultEl.style.display = "none";
    resultEl.className     = "result-box";
    fill.style.width       = "2%";
    document.getElementById(msgId).textContent = "Subiendo archivo…";

    const fd = buildForm(file);
    let jobId;

    try {
      const res  = await fetch(endpoint, { method: "POST", body: fd });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || "Error del servidor");
      jobId = data.job_id;
    } catch (e) {
      showResult(resultEl, "error", "❌ " + e.message);
      btn.disabled = false;
      return;
    }

    pollJob(jobId, fillId, msgId, data => {
      btn.disabled = false;
      if (data.status === "error") {
        showResult(resultEl, "error", "❌ " + (data.msg || "Error en la conversión"));
        return;
      }
      // Detectar caso sin tablas
      if (data.recovered === 0 && endpoint.includes("table")) {
        showResult(resultEl, "warn",
          "⚠️ No se encontraron tablas compatibles.<br>" +
          "Verifica que la tabla haya sido copiada directamente desde Excel.");
        return;
      }
      // Resultado exitoso con descarga
      const n   = data.recovered ? ` — ${data.recovered} tabla(s) recuperada(s)` : "";
      const html =
        `<div>✅ Proceso completado${n}</div>` +
        `<a class="btn-download" href="/api/download/${jobId}">⬇️ Descargar archivo</a>`;
      showResult(resultEl, "ok", html);
    });
  });
}

function showResult(el, type, html) {
  el.innerHTML       = html;
  el.className       = "result-box " + type;
  el.style.display   = "block";
}

// ── FASE 1 ────────────────────────────────────────────────────────────────────
setupRun("run1", "prog1-wrap", "prog1-fill", "prog1-msg", "result1", getFile1,
  file => {
    const fd = new FormData();
    fd.append("file", file);
    return fd;
  },
  "/api/fase1"
);

// ── FASE 2 ────────────────────────────────────────────────────────────────────
setupRun("run2", "prog2-wrap", "prog2-fill", "prog2-msg", "result2", getFile2,
  file => {
    const fd = new FormData();
    fd.append("file", file);
    fd.append("dpi",             document.querySelector('input[name="dpi"]:checked')?.value || "150");
    fd.append("include_text",    document.getElementById("include_text").checked ? "1" : "0");
    fd.append("text_mode",       document.querySelector('input[name="text_mode"]:checked')?.value || "paragraphs");
    fd.append("composition_mode",document.querySelector('input[name="comp"]:checked')?.value || "background");
    fd.append("preserve_line_breaks", document.getElementById("preserve_breaks").checked ? "1" : "0");
    return fd;
  },
  "/api/fase2"
);

// ── PDF TABLE RECOVERY ────────────────────────────────────────────────────────
setupRun("run3", "prog3-wrap", "prog3-fill", "prog3-msg", "result3", getFile3,
  file => {
    const fd = new FormData();
    fd.append("file",         file);
    fd.append("gap_tol",      document.querySelector('input[name="gap"]:checked')?.value  || "18");
    fd.append("max_font_size",document.querySelector('input[name="font_size"]:checked')?.value || "16");
    return fd;
  },
  "/api/table_recovery"
);
