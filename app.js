// Coverage Survey Analysis Tool — static web app driver.
// Loads Pyodide, imports analysis.py, runs generate_report(xlsx_bytes, config),
// and injects the returned HTML into #report-output.

const $ = (id) => document.getElementById(id);

const ui = {
  statusCard:   $("status-card"),
  statusText:   $("status-text"),
  statusDetail: $("status-detail"),
  progressBar:  $("progress-bar"),
  configCard:   $("config-card"),
  resultCard:   $("result-card"),
  fileInput:    $("file-input"),
  generateBtn:  $("generate-btn"),
  generateHint: $("generate-hint"),
  printBtn:     $("print-btn"),
  resetBtn:     $("reset-btn"),
  reportOut:    $("report-output"),
};

let pyodide = null;
let analysisModule = null;
let uploadedBytes = null;
let uploadedName  = null;

function setStatus(text, detail = "", progress = null) {
  ui.statusText.textContent = text;
  ui.statusDetail.textContent = detail;
  if (progress !== null) ui.progressBar.style.width = `${progress}%`;
}

function showError(message) {
  const banner = document.createElement("div");
  banner.className = "error-banner";
  banner.textContent = message;
  ui.statusCard.appendChild(banner);
}

// ── Boot Pyodide and load analysis.py ────────────────────────────────────────
async function bootPyodide() {
  try {
    setStatus("Loading analysis engine…", "Downloading Pyodide runtime (~10 MB).", 5);
    pyodide = await loadPyodide({
      indexURL: "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/",
    });

    setStatus("Loading data libraries…", "pandas, numpy, scipy, matplotlib.", 30);
    await pyodide.loadPackage([
      "numpy", "pandas", "scipy", "matplotlib", "micropip",
    ]);

    setStatus("Loading xlsx support…", "openpyxl via micropip.", 60);
    // openpyxl is pure-Python and not in Pyodide's built-in package index,
    // so we pull it from PyPI through micropip.
    await pyodide.runPythonAsync(`
import micropip
await micropip.install("openpyxl")
`);

    setStatus("Loading report code…", "", 75);
    const analysisSrc = await fetch("analysis.py").then((r) => {
      if (!r.ok) throw new Error(`analysis.py fetch failed (${r.status})`);
      return r.text();
    });
    // Register analysis.py as an importable module.
    pyodide.FS.writeFile("/home/pyodide/analysis.py", analysisSrc);
    analysisModule = pyodide.pyimport("analysis");

    setStatus("Ready.", "Fill in the survey metadata, upload your xlsx, then click Generate.", 100);
    ui.configCard.classList.remove("hidden");

    // Default report-date placeholder to current month.
    const now = new Date();
    const month = now.toLocaleString("en-US", { month: "long", year: "numeric" });
    $("cfg-date").placeholder = `e.g. ${month}`;
    if (!$("cfg-date").value) $("cfg-date").value = month;
  } catch (err) {
    setStatus("Failed to load analysis engine.", String(err), 0);
    showError(String(err));
    throw err;
  }
}

// ── Parse the "District = 0.78" textarea into a {district: fraction} dict ──
function parseReportedCoverage(text) {
  const out = {};
  if (!text || !text.trim()) return null;
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const m = trimmed.match(/^(.+?)\s*[=:]\s*([\d.]+)\s*%?\s*$/);
    if (!m) throw new Error(`Could not parse reported-coverage line: "${trimmed}"`);
    const district = m[1].trim();
    let value = parseFloat(m[2]);
    if (Number.isNaN(value)) throw new Error(`Bad number in: "${trimmed}"`);
    // Accept 78 or 0.78 — normalize to fraction in [0, 1].
    if (value > 1) value /= 100;
    if (value < 0 || value > 1) throw new Error(`Out of range: "${trimmed}"`);
    out[district] = value;
  }
  return Object.keys(out).length ? out : null;
}

function readConfigFromForm() {
  return {
    country:     $("cfg-country").value.trim() || "Country",
    disease:     $("cfg-disease").value.trim() || "Disease",
    drug:        $("cfg-drug").value.trim()    || "Ivermectin (IVM)",
    drug_code:   $("cfg-drug-code").value      || "ivm",
    mda_round:   parseInt($("cfg-mda").value, 10) || 1,
    report_date: $("cfg-date").value.trim() || "",
    sheet_name:  $("cfg-sheet").value.trim() || "data",
    thresh_epi:  parseInt($("cfg-epi").value, 10) || 65,
    thresh_thera:parseInt($("cfg-thera").value, 10) || 80,
    reported_coverage: parseReportedCoverage($("cfg-reported").value),
  };
}

// ── File upload handler ────────────────────────────────────────────────────
ui.fileInput.addEventListener("change", async (ev) => {
  const file = ev.target.files && ev.target.files[0];
  if (!file) {
    uploadedBytes = null;
    uploadedName  = null;
    ui.generateBtn.disabled = true;
    ui.generateHint.textContent = "Upload an xlsx file to enable.";
    return;
  }
  uploadedBytes = new Uint8Array(await file.arrayBuffer());
  uploadedName  = file.name;
  ui.generateBtn.disabled = false;
  ui.generateHint.textContent = `Loaded ${file.name} (${(file.size / 1024).toFixed(0)} KB).`;
});

// ── Generate button ────────────────────────────────────────────────────────
ui.generateBtn.addEventListener("click", async () => {
  if (!uploadedBytes) return;
  if (!analysisModule) {
    showError("Analysis engine not ready yet.");
    return;
  }

  ui.generateBtn.disabled = true;
  ui.generateBtn.textContent = "Generating…";
  ui.reportOut.innerHTML = "";

  try {
    // Marshal the file bytes and config dict over to Python.
    const config = readConfigFromForm();
    const pyBytes = pyodide.toPy(uploadedBytes);
    const pyConfig = pyodide.toPy(config);

    // Run the analysis.
    const start = performance.now();
    const html = analysisModule.generate_report(pyBytes, pyConfig);
    const elapsed = ((performance.now() - start) / 1000).toFixed(1);
    console.log(`Report generated in ${elapsed}s`);

    // analysis.py returns a full HTML *document* with its own embedded styles
    // (.page, .cover, .section-title, .scorecard, etc). Rendering it inline
    // would clobber the app chrome's styles, so we host it inside an iframe.
    // This also makes printing trivial: we call print() on the iframe so only
    // the report goes to the PDF.
    ui.reportOut.innerHTML = "";
    const frame = document.createElement("iframe");
    frame.id = "report-frame";
    frame.setAttribute("title", "Report preview");
    frame.style.width = "100%";
    frame.style.border = "none";
    frame.style.background = "#fff";
    frame.style.minHeight = "400px";
    frame.srcdoc = String(html);
    ui.reportOut.appendChild(frame);

    // Resize the iframe to fit its full content height so the user can scroll
    // through the whole report inline (no inner scrollbar).
    frame.addEventListener("load", () => {
      try {
        const doc = frame.contentDocument;
        const fit = () => {
          const h = Math.max(
            doc.body.scrollHeight,
            doc.documentElement.scrollHeight,
          );
          frame.style.height = h + "px";
        };
        fit();
        // Re-fit on image load (figures arrive as base64 PNGs and may decode
        // after the initial layout).
        doc.querySelectorAll("img").forEach((img) => {
          if (!img.complete) img.addEventListener("load", fit, { once: true });
        });
      } catch (e) {
        console.warn("Could not auto-size iframe:", e);
      }
    });

    ui.resultCard.classList.remove("hidden");
    ui.reportOut.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (err) {
    console.error(err);
    showError(`Report generation failed: ${err.message || err}`);
  } finally {
    ui.generateBtn.disabled = false;
    ui.generateBtn.textContent = "Generate report";
  }
});

// ── Print + reset ──────────────────────────────────────────────────────────
ui.printBtn.addEventListener("click", () => {
  // Print the iframe directly so the saved PDF contains only the report and
  // none of the app chrome (form, buttons, etc).
  const frame = document.getElementById("report-frame");
  if (frame && frame.contentWindow) {
    frame.contentWindow.focus();
    frame.contentWindow.print();
  } else {
    window.print();
  }
});

ui.resetBtn.addEventListener("click", () => {
  ui.reportOut.innerHTML = "";
  ui.resultCard.classList.add("hidden");
  uploadedBytes = null;
  uploadedName = null;
  ui.fileInput.value = "";
  ui.generateBtn.disabled = true;
  ui.generateHint.textContent = "Upload an xlsx file to enable.";
  window.scrollTo({ top: 0, behavior: "smooth" });
});

bootPyodide();
