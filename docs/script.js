let workbook, addressColumn;

function toggleAPIVisibility() {
  const input = document.getElementById("apiKey");
  input.type = input.type === "password" ? "text" : "password";
}

document.getElementById("fileInput").addEventListener("change", handleFile, false);
document.getElementById("processBtn").addEventListener("click", processGeocoding);
document.getElementById("reset-btn").addEventListener("click", resetApp);

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: "array" });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    displayPreview(json);
    populateColumnSelect(json[0]);
  };
  reader.readAsArrayBuffer(file);
}

function displayPreview(data) {
  const preview = data.slice(0, 5).map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`).join("");
  document.getElementById("previewContainer").innerHTML = `<table border="1">${preview}</table>`;
}

function populateColumnSelect(headers) {
  const select = document.getElementById("columnSelect");
  select.innerHTML = "";
  headers.forEach((header, idx) => {
    const option = document.createElement("option");
    option.value = idx;
    option.text = header || `Column ${idx + 1}`;
    select.appendChild(option);
  });
  document.getElementById("columnSelectContainer").style.display = "block";
}

async function processGeocoding() {
  const apiKey = document.getElementById("apiKey").value.trim();
  const colIndex = parseInt(document.getElementById("columnSelect").value);

  if (!apiKey) return alert("Please enter your API key.");
  if (!workbook) return alert("Please upload an Excel file.");

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const header = data[0];
  const rows = data.slice(1);

  header.push("Latitude", "Longitude");

  for (let i = 0; i < rows.length; i++) {
    const address = rows[i][colIndex];
    if (!address) continue;

    try {
      const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
      const result = await response.json();
      if (result.status === "OK") {
        const loc = result.results[0].geometry.location;
        rows[i].push(loc.lat, loc.lng);
      } else {
        rows[i].push("ERROR", "ERROR");
      }
    } catch {
      rows[i].push("ERROR", "ERROR");
    }
  }

  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Geocoded");

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);

  const downloadLink = document.getElementById("downloadLink");
  downloadLink.href = url;
  downloadLink.download = "geocoded_output.xlsx";
  downloadLink.style.display = "block";
  downloadLink.textContent = "⬇ Download Geocoded Excel";
}

function resetApp() {
  document.getElementById("apiKey").value = "";
  document.getElementById("fileInput").value = "";
  document.getElementById("previewContainer").innerHTML = "";
  document.getElementById("columnSelectContainer").style.display = "none";
  document.getElementById("downloadLink").style.display = "none";
  document.getElementById("resultContainer").innerHTML = "";
  document.getElementById("message").innerHTML = "";
  workbook = null;
}
