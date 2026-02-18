async function loadExcelData() {
    const response = await fetch("algae_data.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    return data;
}

function buildTable(data) {
    let html = "<table><tr>";

    // Headers
    Object.keys(data[0]).forEach(key => {
        html += `<th>${key}</th>`;
    });
    html += "</tr>";

    // Rows
    data.forEach(row => {
        html += "<tr>";
        Object.values(row).forEach(val => {
            html += `<td>${val}</td>`;
        });
        html += "</tr>";
    });

    html += "</table>";
    document.getElementById("table-container").innerHTML = html;
}

function buildCellsChart(data) {
    const ctx = document.getElementById("cellsChart");

    const labels = data.map(row => row.SampleDate);
    const values = data.map(row => row.CellsPerML);

    new Chart(ctx, {
        type: "line",
        data: {
            labels: labels,
            datasets: [{
                label: "Cells per mL",
                data: values,
                borderColor: "green",
                tension: 0.2
            }]
        }
    });
}

loadExcelData().then(data => {
    buildTable(data);
    buildCellsChart(data);
});
