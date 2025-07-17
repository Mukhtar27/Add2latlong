let workbook;
let addressData = [];
let selectedColumn = "";
let apiKey = "";

document.getElementById("fileInput").addEventListener("change", handleFile);
document.getElementById("processBtn").addEventListener("click", processAddresses);
document.getElementById("reset-btn").addEventListener("click", () => location.reload());

function toggleAPIVisibility() {
  const apiInput = document.getElementById("apiKey");
  apiInput.type = apiInput.type === "password" ? "text" : "password";
}

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: "array" });

    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    addressData = XLSX.utils.sheet_to_json(firstSheet);

    displayPreview(addressData);
    populateColumnDropdown(Object.keys(addressData[0]));
  };
  reader.readAsArrayBuffer(file);
}

function displayPreview(data) {
  const preview = document.getElementById("previewContainer");
  preview.innerHTML = "<h3>Preview:</h3>";

  const table = document.createElement("table");
  table.className = "preview-table";
  const headers = Object.keys(data[0]);
  const headerRow = table.insertRow();
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });

  data.slice(0, 5).forEach((row) => {
    const tr = table.insertRow();
    headers.forEach((header) => {
      const td = tr.insertCell();
      td.textContent = row[header];
    });
  });

  preview.appendChild(table);
}

function populateColumnDropdown(columns) {
  const dropdown = document.getElementById("columnSelect");
  dropdown.innerHTML = "";
  columns.forEach((col) => {
    const option = document.createElement("option");
    option.value = col;
    option.textContent = col;
    dropdown.appendChild(option);
  });

  document.getElementById("columnSelectContainer").style.display = "block";
}

async function processAddresses() {
  apiKey = document.getElementById("apiKey").value;
  if (!apiKey) {
    alert("Please enter your Google Maps API key.");
    return;
  }

  selectedColumn = document.getElementById("columnSelect").value;
  if (!selectedColumn) {
    alert("Please select the address column.");
    return;
  }

  document.getElementById("message").textContent = "Fetching coordinates...";
  for (let row of addressData) {
    const address = row[selectedColumn];
    const { lat, lng } = await getCoordinates(address);
    row["Latitude"] = lat;
    row["Longitude"] = lng;
  }

  showResults();
  enableDownload();
  document.getElementById("message").textContent = "✅ Coordinates added successfully!";
}

async function getCoordinates(address) {
  try {
    const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`;
    const response = await fetch(url);
    const data = await response.json();

    if (data.status === "OK") {
      const location = data.results[0].geometry.location;
      return { lat: location.lat, lng: location.lng };
    } else {
      return { lat: "", lng: "" };
    }
  } catch {
    return { lat: "", lng: "" };
  }
}

function showResults() {
  const result = document.getElementById("resultContainer");
  result.innerHTML = "<h3>Geocoded Data:</h3>";

  const table = document.createElement("table");
  table.className = "result-table";
  const headers = Object.keys(addressData[0]);
  const headerRow = table.insertRow();
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });

  addressData.forEach((row) => {
    const tr = table.insertRow();
    headers.forEach((header) => {
      const td = tr.insertCell();
      td.textContent = row[header];
    });
  });

  result.appendChild(table);
}

function enableDownload() {
  const worksheet = XLSX.utils.json_to_sheet(addressData);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Geocoded");
  const wbout = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  const downloadLink = document.getElementById("downloadLink");
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.download = "geocoded_addresses.xlsx";
  downloadLink.style.display = "inline-block";
}
