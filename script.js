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
    // Excel serial date → JS Date (UTC base)
    const epoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(epoch.getTime() + d * 86400000);
  }
  const t = String(d ?? "").trim();
  if (!t) return null;
  const dt = new Date(t);
  return isNaN(dt) ? null : dt;
}

// ISO yyyy-mm-dd; change to US format in the commented block below if you prefer.
const fmt = (d) => d ? d.toISOString().slice(0,10) : "";

/*
// US style (MM/DD/YYYY):
// const fmt = (d) => {
//   if (!d) return "";
//   const mm = String(d.getMonth() + 1).padStart(2, "0");
//   const dd = String(d.getDate()).padStart(2, "0");
//   const yyyy = d.getFullYear();
//   return `${mm}/${dd}/${yyyy}`;
// };
*/

function normalize(s) { return String(s || "").toLowerCase().replace(/\s+|_+/g, ""); }

// Try to find a header key by patterns (returns the original header text)
function findKey(keys, patterns) {
  const nk = keys.map(k => ({ k, n: normalize(k) }));
  for (const pat of patterns) {
    const re = new RegExp(pat, "i");
    for (const {k, n} of nk) if (re.test(n)) return k; // normalized
    for (const {k} of nk) if (re.test(k)) return k;    // raw contains
  }
  return null;
}

// Generate a placeholder header for blank cells (e.g., COL_A, COL_B, …)
function placeholderName(i) {
  // A, B, C… Z, AA, AB…
  const letters = (() => {
    let n = i, s = "";
    while (true) {
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26) - 1;
      if (n < 0) break;
    }
    return s;
  })();
  return `COL_${letters}`;
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

  // Score each row for header-likeness using your known columns.
  // Your sheet uses names like: Date_sampled, BioDataTaxonName, Cells_per_mL,
  // NU_per_mL, "Biovolume per Cubic Micrometer (µm³/mL)", PHYLUM, CLASS, ORDER,
  // FAMILY, GENUS, and site/depth fields ("Top", "Middle", "Bottom", "Site 2 T/B", etc.). [1](https://readingareawater-my.sharepoint.com/personal/patrick_bauer_readingareawater_com/_layouts/15/Doc.aspx?sourcedoc=%7B255A45EB-1E69-4B48-9806-49E457349B26%7D&file=algae_data.xlsx&action=default&mobileredirect=true)
  const headerPatterns = [
    "date", "biodatataxonname", "scientific", "taxon",
    "cellsperml", "nuperml", "biovolume",
    "^phylum$", "^class$", "^order$", "^family$", "^genus$",
    "normalizedsite", "site", "location", "depth",
    "algalgroup", "project", "lab"
  ];
  function rowScore(arr) {
    const cells = (arr || []).map(c => normalize(c));
    let score = 0;
    for (const c of cells) {
      for (const pat of headerPatterns) {
        if (new RegExp(pat, "i").test(c)) { score++; break; }
      }
    }
    return score;
  }

  // Pick the row with the highest score; require a minimum to avoid data rows.
  let bestIdx = -1, bestScore = -1;
  rows.forEach((r, i) => {
    const s = rowScore(r);
    if (s > bestScore) { bestScore = s; bestIdx = i; }
  });

  if (bestIdx < 0 || bestScore < 5) {
    throw new Error("Could not locate the header row in the Excel sheet.");
  }

  // Build a clean header: trim, and fill blanks with unique placeholders
  const rawHeader = rows[bestIdx] || [];
  const header = rawHeader.map((h, i) => {
    const name = (h ?? "").toString().trim();
    return name ? name : placeholderName(i);
  });

  // Ensure headers are unique (avoid collisions)
  const seen = new Set();
  for (let i = 0; i < header.length; i++) {
    let name = header[i];
    let j = 1;
    while (seen.has(name)) {
      name = `${header[i]}_${++j}`;
    }
    header[i] = name;
    seen.add(name);
  }

  // Build data objects from subsequent rows; initialize each object with ALL headers
  const dataRows = rows.slice(bestIdx + 1);
  const objects = dataRows
    .filter(r => r && r.some(cell => cell !== null && cell !== "")) // keep rows with any content
    .map(r => {
      const obj = {};
      header.forEach(h => { obj[h] = null; });          // initialize all keys
      header.forEach((h, i) => { obj[h] = r[i] ?? null; }); // fill present cells
      return obj;
    });

  return { header, rows: objects };
}

/***** Table (dates formatted; blanks handled) *****/
function buildTable(data, options = {}) {
  if (!data.rows.length) { $("#table-container").innerHTML = "<p>No rows.</p>"; return; }

  const { header } = data;
  const dateColumns = header.filter(h => /date/i.test(h)); // any header containing "date"

  let html = '<div class="table-wrap"><table><thead><tr>';
  header.forEach(k => html += `<th>${k}</th>`);
  html += '</tr></thead><tbody>';

  data.rows.forEach(row => {
    html += "<tr>";
    header.forEach(k => {
      let val = row[k];

      // Format any header that looks like a date column
      if (dateColumns.includes(k)) {
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
  // Filter by selected site/depth (e.g., "Top", "Site 2 B", etc.). [1](https://readingareawater-my.sharepoint.com/personal/patrick_bauer_readingareawater_com/_layouts/15/Doc.aspx?sourcedoc=%7B255A45EB-1E69-4B48-9806-49E457349B26%7D&file=algae_data.xlsx&action=default&mobileredirect=true)
  let filtered = data.rows.filter(r => (r[siteKey] || "").toString().trim() === selectedSite);

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
    const loaded = await loadExcelRows();
    if (!loaded.rows.length) {
      setStatus("No data rows found in algae_data.xlsx.", true);
      return;
    }

    setStatus("Mapping columns…");
    const header = loaded.header;

    // Detect column names from your sheet (flexible). [1](https://readingareawater-my.sharepoint.com/personal/patrick_bauer_readingareawater_com/_layouts/15/Doc.aspx?sourcedoc=%7B255A45EB-1E69-4B48-9806-49E457349B26%7D&file=algae_data.xlsx&action=default&mobileredirect=true)
    const dateKey   = findKey(header, ["^date.?sampled$", "date"]);
    const taxonKey  = findKey(header, ["^biodatataxonname$", "scientific", "taxon"]);
    const cellsKey  = findKey(header, ["^cells.?per.?ml$"]);
    const nuKey     = findKey(header, ["^nu.?per.?ml$"]);
    const bioKey    = findKey(header, ["biovolume"]); // matches the long Biovolume header
    const siteKey   = findKey(header, ["^normalizedsite$", "site", "location", "depth"]);
    const genusKey  = findKey(header, ["^genus$"]);

    if (!dateKey || !taxonKey || !cellsKey || !siteKey) {
      throw new Error("Missing required columns (need at least Date, Taxon, Cells_per_mL, Site).");
    }

    // Normalize a "Biovolume" property for UI use (non-destructive)
    const data = {
      header,
      rows: loaded.rows.map(r => {
        const obj = { ...r };
        if (bioKey && obj.Biovolume == null) obj.Biovolume = obj[bioKey];
        return obj;
      })
    };

    // Populate site selector
    const sites = Array.from(new Set(data.rows
      .map(r => (r[siteKey] || "").toString().trim())
      .filter(Boolean))).sort();
    const siteSel = $("#siteSelect");
    siteSel.innerHTML = sites.map(s => `<option>${s}</option>`).join("");

    // Disable metrics that aren't present
    const metricSel = $("#metricSelect");
    if (!nuKey)  metricSel.querySelector('option[value="NU_per_mL"]').disabled = true;
    if (!bioKey) metricSel.querySelector('option[value="Biovolume"]').disabled = true;

    // Initial chart/table
    buildChart(data, siteKey, dateKey, cellsKey, genusKey, sites[0], "Cells_per_mL", "");
    buildTable(data, { /* formats any 'date' headers automatically */ });

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
``
