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

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ── Dynamic reported-coverage inputs ─────────────────────────────────────────
// Called after a file is uploaded; uses SheetJS (loaded from CDN) to read
// the configured sheet and extract unique p_district values, then builds
// one number input per district so the user can enter reported coverage.

function buildReportedCoverageInputs(districts) {
  const container = $("reported-coverage-inputs");
  if (!districts || districts.length === 0) {
    container.innerHTML =
      '<p class="muted small" style="font-style:italic">' +
      "No districts detected — check the sheet name and re-upload.</p>";
    return;
  }

  const grid = document.createElement("div");
  grid.className = "grid";

  for (const d of districts) {
    const label = document.createElement("label");
    label.innerHTML =
      `${escapeHtml(d)} <span class="muted">(reported %)</span>` +
      `<input type="number" class="rc-input" data-district="${escapeHtml(d)}"` +
      ` min="0" max="100" step="0.1" placeholder="e.g. 78.5" />`;
    grid.appendChild(label);
  }

  container.innerHTML = "";
  container.appendChild(grid);
}

function detectDistrictsFromXlsx(bytes) {
  try {
    const wb = XLSX.read(bytes, { type: "array" });
    const sheetName = $("cfg-sheet").value.trim() || "data";
    const ws = wb.Sheets[sheetName];
    if (!ws) {
      buildReportedCoverageInputs([]);
      return;
    }
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
    const seen = new Set();
    const districts = [];
    for (const row of rows) {
      const d = row.p_district;
      if (d && String(d).trim() && !seen.has(String(d).trim())) {
        seen.add(String(d).trim());
        districts.push(String(d).trim());
      }
    }
    districts.sort();
    buildReportedCoverageInputs(districts);
  } catch (e) {
    console.warn("District detection failed:", e);
    buildReportedCoverageInputs([]);
  }
}

// Re-detect districts if the user changes the sheet name after uploading.
$("cfg-sheet").addEventListener("input", () => {
  if (uploadedBytes) detectDistrictsFromXlsx(uploadedBytes);
});

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
    await pyodide.runPythonAsync(`
import micropip
await micropip.install("openpyxl")
`);

    setStatus("Loading report code…", "", 75);
    const analysisSrc = await fetch("analysis.py").then((r) => {
      if (!r.ok) throw new Error(`analysis.py fetch failed (${r.status})`);
      return r.text();
    });
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

// ── Config collection ─────────────────────────────────────────────────────────
function readConfigFromForm() {
  // Collect per-district reported coverage from the dynamic inputs.
  const rcInputs = document.querySelectorAll("#reported-coverage-inputs .rc-input");
  let reported_coverage = null;
  if (rcInputs.length > 0) {
    const out = {};
    for (const inp of rcInputs) {
      const district = inp.dataset.district;
      const raw = inp.value.trim();
      if (!raw) continue;
      let val = parseFloat(raw);
      if (isNaN(val)) continue;
      if (val > 1) val /= 100;   // accept 78 or 0.78
      if (val < 0 || val > 1) continue;
      out[district] = val;
    }
    if (Object.keys(out).length > 0) reported_coverage = out;
  }

  return {
    country:           $("cfg-country").value.trim() || "Country",
    disease:           $("cfg-disease").value.trim() || "Disease",
    drug:              $("cfg-drug").value.trim()    || "Ivermectin (IVM)",
    drug_code:         $("cfg-drug-code").value      || "ivm",
    mda_round:         parseInt($("cfg-mda").value, 10) || 1,
    report_date:       $("cfg-date").value.trim() || "",
    sheet_name:        $("cfg-sheet").value.trim() || "data",
    thresh_epi:        parseInt($("cfg-epi").value, 10) || 65,
    thresh_thera:      parseInt($("cfg-thera").value, 10) || 80,
    reported_coverage,
  };
}

// ── File upload handler ────────────────────────────────────────────────────
ui.fileInput.addEventListener("change", async (ev) => {
  const file = ev.target.files && ev.target.files[0];
  if (!file) {
    uploadedBytes = null;
    uploadedName  = null;
    ui.generateBtn.disabled = true;
    $("generate-hint").textContent = "Upload an xlsx file to enable.";
    $("reported-coverage-inputs").innerHTML =
      '<p class="muted small" style="font-style:italic">Waiting for data file…</p>';
    return;
  }
  uploadedBytes = new Uint8Array(await file.arrayBuffer());
  uploadedName  = file.name;
  ui.generateBtn.disabled = false;
  $("generate-hint").textContent =
    `Loaded ${file.name} (${(file.size / 1024).toFixed(0)} KB).`;

  // Detect IUs/districts from the xlsx and populate the reported-coverage form.
  detectDistrictsFromXlsx(uploadedBytes);
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
    const config = readConfigFromForm();
    const pyBytes = pyodide.toPy(uploadedBytes);
    const pyConfig = pyodide.toPy(config);

    const start = performance.now();
    const html = analysisModule.generate_report(pyBytes, pyConfig);
    const elapsed = ((performance.now() - start) / 1000).toFixed(1);
    console.log(`Report generated in ${elapsed}s`);

    // Render the report inside an iframe so its embedded styles don't clash
    // with the app chrome, and so Save as PDF prints only the report.
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
  $("generate-hint").textContent = "Upload an xlsx file to enable.";
  $("reported-coverage-inputs").innerHTML =
    '<p class="muted small" style="font-style:italic">Waiting for data file…</p>';
  window.scrollTo({ top: 0, behavior: "smooth" });
});

bootPyodide();
