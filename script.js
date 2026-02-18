/***** Utilities *****/
const $ = (sel) => document.querySelector(sel);
const setStatus = (msg, isError=false) => {
  const el = $("#status");
  el.textContent = msg;
  el.className = isError ? "status error" : "status";
};

// Convert Excel serial / Date / string → Date
function toDate(d) {
  if (d instanceof Date) return d;
  if (typeof d === "number" && isFinite(d)) {
    // Excel serial date → JS Date (UTC)
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + d * 86400000);
  }
  const t = String(d ?? "").trim();
  if (!t) return null;
  const dt = new Date(t);
  return isNaN(dt) ? null : dt;
}

// Format date (change to US style below if you prefer)
const fmt = (d) => d ? d.toISOString().slice(0,10) : "";

/*
// US style:
// const fmt = (d) => {
//   if (!d) return "";
//   const mm = String(d.getMonth() + 1).padStart(2, "0");
//   const dd = String(d.getDate()).padStart(2, "0");
//   const yyyy = d.getFullYear();
//   return `${mm}/${dd}/${yyyy}`;
// };
*/

// Find a key in a header list by pattern (ignore spaces/underscores)
function findKey(keys, patterns) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+|_+/g, "");
  const nk = keys.map(k => ({ k, n: norm(k) }));
  for (const pat of patterns) {
    const re = new RegExp(pat, "i");
    for (const {k, n} of nk) if (re.test(n)) return k; // normalized match
    for (const k of keys) if (re.test(k)) return k;    // raw contains
  }
  return null;
}

/***** Load Excel and robustly detect the header row *****/
async function loadExcelRows() {
  const res = await fetch("algae_data.xlsx");
  if (!res.ok) throw new Error(`Failed to fetch algae_data.xlsx (HTTP ${res.status})`);
  const buf = await res.arrayBuffer();

  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  // Read as 2D array; keep blanks as null
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!rows.length) return [];

  // Skip the title/blank rows and detect the true header line:
  // require a row containing Date + Taxon + Cells per mL and enough non-empty cells
  // (matches your workbook's header names like Date_sampled, BioDataTaxonName, Cells_per_mL, long Biovolume name, etc.). [1](https://readingareawater-my.sharepoint.com/personal/patrick_bauer_readingareawater_com/_layouts/15/Doc.aspx?sourcedoc=%7B255A45EB-1E69-4B48-9806-49E457349B26%7D&file=algae_data.xlsx&action=default&mobileredirect=true)
  const headerRowIdx = rows.findIndex(r => {
    if (!r) return false;
    const cells = r.map(c => (c ?? "").toString().toLowerCase());
    const hasDate  = cells.some(c => /date/.test(c));
    const hasTaxon = cells.some(c => /(biodatataxonname|scientific|taxon)/.test(c));
    const hasCells = cells.some(c => /cells\s*per\s*ml|cellsperml/.test(c));
    const nonEmpty = cells.filter(c => c.trim() !== "").length;
    return hasDate && hasTaxon && hasCells && nonEmpty >= 6;
  });

  if (headerRowIdx < 0) {
    throw new Error("Could not locate the header row in the Excel sheet.");
  }

  // Build the header from that row
  const header = (rows[headerRowIdx] || []).map(h => (h ?? "").toString().trim());

  // Build data objects from subsequent rows, skipping fully empty ones
  const dataRows = rows.slice(headerRowIdx + 1);
  const objects = dataRows
    .filter(r => r && r.some(cell => cell !== null && cell !== "")) // keep rows with any content
    .map(r => {
      const obj = {};
      header.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });

  return objects;
}

/***** Table (dates formatted; blanks handled) *****/
function buildTable(data, options = {}) {
  if (!data.length) { $("#table-container").innerHTML = "<p>No rows.</p>"; return; }

  const { dateKey } = options; // detected date header name
  const keys = Object.keys(data[0]);

  let html = '<div class="table-wrap"><table><thead><tr>';
  keys.forEach(k => html += `<th>${k}</th>`);
  html += '</tr></thead><tbody>';

  data.forEach(row => {
    html += "<tr>";
    keys.forEach(k => {
      let val = row[k];

      // Date column → format (also converts Excel serials)
      if (dateKey && k === dateKey) {
        const d = toDate(val);
        val = d ? fmt(d) : (val ?? "");
      }

      // Clean undefined/null; trim strings
      if (val == null) val = "";
      if (typeof val === "string") val = val.trim();

      html += `<td>${val}</td>`;
    });
    html += "</tr>";
  });

  html += "</tbody></table></div>";
  $("#table-container").innerHTML = html;
}

/***** Chart *****/
let chart;
function buildChart(data, siteKey, dateKey, metricKey, genusKey, selectedSite, selectedMetric, genusFilter) {
  // Filter by selected site/depth (e.g., "Top", "Site 2 B", etc.) from your data. [1](https://readingareawater-my.sharepoint.com/personal/patrick_bauer_readingareawater_com/_layouts/15/Doc.aspx?sourcedoc=%7B255A45EB-1E69-4B48-9806-49E457349B26%7D&file=algae_data.xlsx&action=default&mobileredirect=true)
  let filtered = data.filter(r => (r[siteKey] || "").toString().trim() === selectedSite);

  // Optional genus filter (contains, case-insensitive)
  if (genusFilter) {
    const g = genusFilter.trim().toLowerCase();
    filtered = filtered.filter(r => (r[genusKey] || "").toString().toLowerCase().includes(g));
  }

  // Aggregate by date (sum the metric across taxa)
  const byDate = new Map();
  filtered.forEach(r => {
    const d = toDate(r[dateKey]); if (!d) return;
    const key = fmt(d);
    const raw = r[metricKey];
    const num = (typeof raw === "number") ? raw : parseFloat(String(raw ?? "").replace(/,/g,""));
    const val = isFinite(num) ? num : 0;
    byDate.set(key, (byDate.get(key) || 0) + val);
  });

  const labels = Array.from(byDate.keys()).sort();
  const values = labels.map(l => byDate.get(l));

  const ctx = $("#cellsChart").getContext("2d");
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [{
        label: selectedMetric === "Cells_per_mL" ? "Cells per mL"
             : selectedMetric === "NU_per_mL" ? "NU per mL"
             : "Biovolume (µm³/mL)",
        data: values,
        borderColor: "#2e7d32",
        backgroundColor: "rgba(46,125,50,0.15)",
        tension: 0.2,
        pointRadius: 3
      }]
    },
    options: {
      responsive: true,
      scales: {
        x: { title: { text: "Sample date", display: true } },
        y: { title: { text: "Value", display: true }, beginAtZero: true }
      },
      plugins: { legend: { display: true }, tooltip: { mode: "index", intersect: false } }
    }
  });
}

/***** Main *****/
(async function main() {
  try {
    setStatus("Loading Excel…");
    const raw = await loadExcelRows();
    if (!raw.length) { setStatus("No data rows found in algae_data.xlsx.", true); return; }

    setStatus("Mapping columns…");
    const keys = Object.keys(raw[0]);

    // Detect column names from your sheet (flexible to slight header name differences). [1](https://readingareawater-my.sharepoint.com/personal/patrick_bauer_readingareawater_com/_layouts/15/Doc.aspx?sourcedoc=%7B255A45EB-1E69-4B48-9806-49E457349B26%7D&file=algae_data.xlsx&action=default&mobileredirect=true)
    const dateKey   = findKey(keys, ["^date.?sampled$", "date"]);
    const taxonKey  = findKey(keys, ["^biodatataxonname$", "scientific", "taxon"]);
    const cellsKey  = findKey(keys, ["^cells.?per.?ml$"]);
    const nuKey     = findKey(keys, ["^nu.?per.?ml$"]);
    const bioKey    = findKey(keys, ["biovolume"]); // matches "Biovolume per Cubic Micrometer (µm³/mL)"
    const siteKey   = findKey(keys, ["^normalizedsite$", "site", "location", "depth"]);
    const genusKey  = findKey(keys, ["^genus$"]);

    if (!dateKey || !taxonKey || !cellsKey || !siteKey) {
      throw new Error("Missing required columns (need at least Date, Taxon, Cells_per_mL, Site).");
    }

    // Normalize a "Biovolume" property for UI use
    const data = raw.map(r => {
      const obj = { ...r };
      if (bioKey && !obj.Biovolume) obj.Biovolume = obj[bioKey];
      return obj;
    });

    // Populate site selector
    const sites = Array.from(new Set(data.map(r => (r[siteKey] || "").toString().trim()).filter(Boolean))).sort();
    const siteSel = $("#siteSelect");
    siteSel.innerHTML = sites.map(s => `<option>${s}</option>`).join("");

    // Disable metrics that aren't present
    const metricSel = $("#metricSelect");
    if (!nuKey)  metricSel.querySelector('option[value="NU_per_mL"]').disabled = true;
    if (!bioKey) metricSel.querySelector('option[value="Biovolume"]').disabled = true;

    // Initial chart/table
    buildChart(data, siteKey, dateKey, cellsKey, genusKey, sites[0], "Cells_per_mL", "");
    buildTable(data, { dateKey });   // ← passes dateKey so table formats dates

    setStatus("Loaded ✔");

    // Apply button
    $("#applyBtn").addEventListener("click", () => {
      const selectedSite   = siteSel.value;
      const selectedMetric = $("#metricSelect").value;
      const genusFilter    = $("#genusFilter").value;

      const metricKey =
        selectedMetric === "NU_per_mL" ? (nuKey || cellsKey) :
        selectedMetric === "Biovolume" ? "Biovolume" :
        cellsKey;

      buildChart(data, siteKey, dateKey, metricKey, genusKey, selectedSite, selectedMetric, genusFilter);
    });

  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err.message}`, true);
  }
})();
