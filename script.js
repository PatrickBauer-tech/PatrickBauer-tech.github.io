/***** Small helpers *****/
const $ = (sel) => document.querySelector(sel);
const setStatus = (msg, isError=false) => {
  const el = $("#status");
  el.textContent = msg;
  el.className = isError ? "status error" : "status";
};

// Convert Excel serials / Dates / strings → JS Date
function toDate(d) {
  if (d instanceof Date) return d;
  if (typeof d === "number" && isFinite(d)) {
    // Excel serial origin (Dec 30, 1899)
    const origin = new Date(Date.UTC(1899, 11, 30));
    return new Date(origin.getTime() + d * 86400000);
  }
  const t = String(d ?? "").trim();
  if (!t) return null;
  const dt = new Date(t);
  return isNaN(dt) ? null : dt;
}

// Format date as YYYY-MM-DD (change to US format below if desired)
const fmt = (d) => d ? d.toISOString().slice(0,10) : "";

/*
// US style MM/DD/YYYY:
const fmt = (d) => {
  if (!d) return "";
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
};
*/

// Fixed column order based on your new header row
// (keeps the table aligned and predictable)
const COLUMNS = [
  "Lab2_ID",
  "Date_sampled",
  "BioDataTaxonName",
  "ALGALGROUP",
  "Cells_per_mL",
  "NU_per_mL",
  "Biovolume per Cubic Micrometer (µm³/mL)",
  "PHYLUM",
  "CLASS",
  "ORDER",
  "FAMILY",
  "GENUS",
];

// Load the Excel (first sheet) → array of objects with those headers
async function loadExcel() {
  const res = await fetch("algae_data.xlsx", { cache: "no-store" });
  if (!res.ok) throw new Error(`Failed to fetch algae_data.xlsx (HTTP ${res.status})`);
  const buf = await res.arrayBuffer();

  // Ask SheetJS to keep dates as JS Date when possible
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];

  // Use the first row as the header; keep empty cells as null
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

  // Ensure only our expected columns, in the expected order
  const data = rows.map(r => {
    const o = {};
    COLUMNS.forEach(k => { o[k] = r[k] ?? null; });
    return o;
  });

  return data;
}

/***** Table rendering (dates formatted) *****/
function buildTable(data) {
  if (!data.length) { $("#table-container").innerHTML = "<p>No rows.</p>"; return; }

  let html = '<div class="table-wrap"><table><thead><tr>';
  COLUMNS.forEach(k => { html += `<th>${k}</th>`; });
  html += '</tr></thead><tbody>';

  data.forEach(row => {
    html += "<tr>";
    COLUMNS.forEach(k => {
      let val = row[k];

      // Format Date_sampled
      if (k === "Date_sampled") {
        const d = toDate(val);
        val = d ? fmt(d) : (val ?? "");
      }

      // Clean up null/undefined and trim strings
      if (val == null) val = "";
      if (typeof val === "string") val = val.trim();

      html += `<td>${val}</td>`;
    });
    html += "</tr>";
  });

  html += "</tbody></table></div>";
  $("#table-container").innerHTML = html;
}

/***** Chart rendering *****/
let chart;
function buildChart(data, selectedSite, selectedMetric, genusFilter) {
  const siteKey   = "Lab2_ID";
  const dateKey   = "Date_sampled";
  const genusKey  = "GENUS";
  const metricKey = selectedMetric; // one of the three metric fields

  // Filter by site
  let filtered = data.filter(r => (r[siteKey] || "").toString().trim() === selectedSite);

  // Optional genus filter (contains, case-insensitive)
  if (genusFilter && genusFilter.trim()) {
    const g = genusFilter.trim().toLowerCase();
    filtered = filtered.filter(r => (r[genusKey] || "").toString().toLowerCase().includes(g));
  }

  // Aggregate by date (sum of the metric across taxa for that date/site)
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
        label:
          selectedMetric === "Cells_per_mL" ? "Cells per mL" :
          selectedMetric === "NU_per_mL" ? "NU per mL" :
          "Biovolume (µm³/mL)",
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
      plugins: { legend: { display: true } }
    }
  });
}

/***** Main *****/
(async function main() {
  try {
    setStatus("Loading Excel…");
    const data = await loadExcel();

    if (!data.length) {
      setStatus("No data rows found in algae_data.xlsx.", true);
      return;
    }

    // Populate site selector from Lab2_ID
    const sites = Array.from(new Set(
      data.map(r => (r.Lab2_ID || "").toString().trim()).filter(Boolean)
    )).sort();
    const siteSel = $("#siteSelect");
    siteSel.innerHTML = sites.map(s => `<option>${s}</option>`).join("");

    // Initial chart + table
    const metricSel = $("#metricSelect");
    buildChart(data, sites[0], metricSel.value, "");
    buildTable(data);
    setStatus("Loaded ✔");

    // Apply filters
    $("#applyBtn").addEventListener("click", () => {
      const selectedSite   = siteSel.value;
      const selectedMetric = metricSel.value;
      const genusFilter    = $("#genusFilter").value;
      buildChart(data, selectedSite, selectedMetric, genusFilter);
    });

  } catch (err) {
    console.error(err);
    setStatus(`Error: ${err.message}`, true);
  }
})();
