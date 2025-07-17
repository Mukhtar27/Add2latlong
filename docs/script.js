let parsedData = [];
let selectedColumn = '';

document.getElementById("fileInput").addEventListener("change", handleFileUpload);
document.getElementById("processBtn").addEventListener("click", geocodeAddresses);

function toggleAPIVisibility() {
  const input = document.getElementById("apiKey");
  input.type = input.type === "password" ? "text" : "password";
}

function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  document.getElementById("fileInfo").innerText = `📄 ${file.name}`;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    parsedData = XLSX.utils.sheet_to_json(sheet);

    renderPreview(parsedData);
    populateColumnDropdown(Object.keys(parsedData[0] || {}));
  };
  reader.readAsArrayBuffer(file);
}

function renderPreview(data) {
  const preview = document.getElementById("previewContainer");
  preview.innerHTML = "<h3>Preview of uploaded file:</h3>";

  if (data.length === 0) return;

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  const headerRow = document.createElement("tr");
  Object.keys(data[0]).forEach(col => {
    const th = document.createElement("th");
    th.innerText = col;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  data.slice(0, 5).forEach(row => {
    const tr = document.createElement("tr");
    Object.values(row).forEach(cell => {
      const td = document.createElement("td");
      td.innerText = cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  preview.appendChild(table);
}

function populateColumnDropdown(columns) {
  const select = document.getElementById("columnSelect");
  select.innerHTML = "";
  columns.forEach(col => {
    const option = document.createElement("option");
    option.value = col;
    option.textContent = col;
    select.appendChild(option);
  });
  document.getElementById("columnSelectContainer").style.display = "block";
}

async function geocodeAddresses() {
  const apiKey = document.getElementById("apiKey").value.trim();
  selectedColumn = document.getElementById("columnSelect").value;

  if (!apiKey) {
    alert("Please enter your Google Maps API Key.");
    return;
  }

  const updatedData = await Promise.all(parsedData.map(async (row) => {
    const address = row[selectedColumn];
    if (!address) return { ...row, Latitude: "", Longitude: "" };

    const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`;
    try {
      const response = await fetch(url);
      const data = await response.json();
      if (data.status === "OK") {
        const location = data.results[0].geometry.location;
        return { ...row, Latitude: location.lat, Longitude: location.lng };
      }
    } catch (err) {
      console.error("Error fetching geocode:", err);
    }
    return { ...row, Latitude: "", Longitude: "" };
  }));

  showSuccessTable(updatedData);
  createDownload(updatedData);
}

function showSuccessTable(data) {
  document.getElementById("message").textContent = "✅ Coordinates added successfully!";
  const result = document.getElementById("resultContainer");
  result.innerHTML = "";

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  const headerRow = document.createElement("tr");
  Object.keys(data[0]).forEach(col => {
    const th = document.createElement("th");
    th.innerText = col;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  data.forEach(row => {
    const tr = document.createElement("tr");
    Object.values(row).forEach(cell => {
      const td = document.createElement("td");
      td.innerText = cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  result.appendChild(table);
}

function createDownload(data) {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Geocoded");

  const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  const downloadLink = document.getElementById("downloadLink");
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.download = "geocoded_addresses.xlsx";
  downloadLink.style.display = "block";
}
