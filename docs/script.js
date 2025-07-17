let workbookData = null;
let selectedSheetName = null;
let addressColumn = null;

document.getElementById("fileUpload").addEventListener("change", handleFile, false);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    selectedSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[selectedSheetName];
    workbookData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    const columns = Object.keys(workbookData[0]);
    const select = document.getElementById("columnSelect");
    select.innerHTML = columns.map(col => `<option value="${col}">${col}</option>`).join('');
  };
  reader.readAsArrayBuffer(e.target.files[0]);
}

async function geocodeAddress(address, apiKey) {
  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`;
  try {
    const response = await fetch(url);
    const data = await response.json();
    if (data.status === "OK") {
      const location = data.results[0].geometry.location;
      return [location.lat, location.lng];
    }
  } catch (error) {
    console.error("Geocoding error:", error);
  }
  return [null, null];
}

async function processFile() {
  const apiKey = document.getElementById("apiKey").value.trim();
  addressColumn = document.getElementById("columnSelect").value;

  if (!apiKey || !workbookData || !addressColumn) {
    alert("Make sure API Key, File, and Address Column are selected.");
    return;
  }

  const outputDiv = document.getElementById("output");
  outputDiv.innerHTML = "⏳ Processing... Please wait.";

  for (let row of workbookData) {
    const address = row[addressColumn];
    if (address) {
      const [lat, lon] = await geocodeAddress(address, apiKey);
      row["Latitude"] = lat;
      row["Longitude"] = lon;
    }
  }

  outputDiv.innerHTML = "✅ Coordinates fetched!";
  const ws = XLSX.utils.json_to_sheet(workbookData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Geocoded");

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  const downloadLink = document.getElementById("downloadLink");
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.style.display = "inline-block";
  downloadLink.textContent = "⬇ Download Result Excel";
}
